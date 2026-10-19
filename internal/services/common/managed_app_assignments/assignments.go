// Package managedappassignments adds assignment support to the Intune MAM
// (app protection) policy resources, which the Graph SDK only exposes through
// a hand-built request to /deviceAppManagement/managedAppPolicies/{id}/assign.
package managedappassignments

import (
	"context"
	"fmt"
	"regexp"

	"github.com/deploymenttheory/terraform-provider-microsoft365/internal/constants"
	"github.com/hashicorp/terraform-plugin-framework-validators/stringvalidator"
	"github.com/hashicorp/terraform-plugin-framework/attr"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/planmodifier"
	"github.com/hashicorp/terraform-plugin-framework/resource/schema/setplanmodifier"
	"github.com/hashicorp/terraform-plugin-framework/schema/validator"
	"github.com/hashicorp/terraform-plugin-framework/types"
	abstractions "github.com/microsoft/kiota-abstractions-go"
	msgraphbetasdk "github.com/microsoftgraph/msgraph-beta-sdk-go"
	"github.com/microsoftgraph/msgraph-beta-sdk-go/deviceappmanagement"
	graphmodels "github.com/microsoftgraph/msgraph-beta-sdk-go/models"
	"github.com/microsoftgraph/msgraph-beta-sdk-go/models/odataerrors"
)

const (
	TypeAllLicensedUsers = "allLicensedUsersAssignmentTarget"
	TypeGroup            = "groupAssignmentTarget"
	TypeExclusionGroup   = "exclusionGroupAssignmentTarget"

	assignURLTemplate = "{+baseurl}/deviceAppManagement/managedAppPolicies/{managedAppPolicyId}/assign"
)

// AssignmentModel is one element of the `assignments` set.
type AssignmentModel struct {
	Type    types.String `tfsdk:"type"`
	GroupId types.String `tfsdk:"group_id"`
}

// AttrTypes is the object type of one assignment, used to build types.Set values.
var AttrTypes = map[string]attr.Type{
	"type":     types.StringType,
	"group_id": types.StringType,
}

// Schema returns the `assignments` attribute shared by the Android and iOS policies.
func Schema() schema.SetNestedAttribute {
	return schema.SetNestedAttribute{
		Optional: true,
		Computed: true,
		PlanModifiers: []planmodifier.Set{
			setplanmodifier.UseStateForUnknown(),
		},
		MarkdownDescription: "Assignments for the app protection policy. Omit to leave the policy unassigned. " +
			"Managing this attribute replaces the full assignment list on the policy. " +
			"App protection policies target users/groups, never devices.",
		NestedObject: schema.NestedAttributeObject{
			Attributes: map[string]schema.Attribute{
				"type": schema.StringAttribute{
					Required: true,
					MarkdownDescription: "One of 'allLicensedUsersAssignmentTarget' (All users), " +
						"'groupAssignmentTarget' or 'exclusionGroupAssignmentTarget'.",
					Validators: []validator.String{
						stringvalidator.OneOf(TypeAllLicensedUsers, TypeGroup, TypeExclusionGroup),
					},
				},
				"group_id": schema.StringAttribute{
					Optional:            true,
					MarkdownDescription: "Entra ID group ID. Required for group and exclusion group targets; must be omitted for all licensed users.",
					Validators: []validator.String{
						stringvalidator.RegexMatches(
							regexp.MustCompile(constants.GuidRegex),
							"must be a valid GUID in the format 00000000-0000-0000-0000-000000000000",
						),
					},
				},
			},
		},
	}
}

// Construct converts the Terraform set into Graph assignment objects.
func Construct(ctx context.Context, set types.Set) ([]graphmodels.TargetedManagedAppPolicyAssignmentable, error) {
	result := make([]graphmodels.TargetedManagedAppPolicyAssignmentable, 0)
	if set.IsNull() || set.IsUnknown() {
		return result, nil
	}

	var items []AssignmentModel
	if diags := set.ElementsAs(ctx, &items, false); diags.HasError() {
		return nil, fmt.Errorf("reading assignments: %v", diags)
	}

	for _, item := range items {
		a := graphmodels.NewTargetedManagedAppPolicyAssignment()

		switch item.Type.ValueString() {
		case TypeAllLicensedUsers:
			if !item.GroupId.IsNull() && item.GroupId.ValueString() != "" {
				return nil, fmt.Errorf("group_id must not be set when type is %s", TypeAllLicensedUsers)
			}
			t := graphmodels.NewAllLicensedUsersAssignmentTarget()
			odata := "#microsoft.graph." + TypeAllLicensedUsers
			t.SetOdataType(&odata)
			a.SetTarget(t)
		case TypeGroup:
			if item.GroupId.IsNull() || item.GroupId.ValueString() == "" {
				return nil, fmt.Errorf("group_id is required when type is %s", TypeGroup)
			}
			t := graphmodels.NewGroupAssignmentTarget()
			odata := "#microsoft.graph." + TypeGroup
			t.SetOdataType(&odata)
			gid := item.GroupId.ValueString()
			t.SetGroupId(&gid)
			a.SetTarget(t)
		case TypeExclusionGroup:
			if item.GroupId.IsNull() || item.GroupId.ValueString() == "" {
				return nil, fmt.Errorf("group_id is required when type is %s", TypeExclusionGroup)
			}
			t := graphmodels.NewExclusionGroupAssignmentTarget()
			odata := "#microsoft.graph." + TypeExclusionGroup
			t.SetOdataType(&odata)
			gid := item.GroupId.ValueString()
			t.SetGroupId(&gid)
			a.SetTarget(t)
		default:
			return nil, fmt.Errorf("unsupported assignment type %q", item.Type.ValueString())
		}

		result = append(result, a)
	}
	return result, nil
}

// MapToTerraform converts Graph assignments (GET .../assignments) into a types.Set.
func MapToTerraform(assignments []graphmodels.TargetedManagedAppPolicyAssignmentable) (types.Set, error) {
	objType := types.ObjectType{AttrTypes: AttrTypes}
	values := make([]attr.Value, 0, len(assignments))

	for _, a := range assignments {
		target := a.GetTarget()
		if target == nil || target.GetOdataType() == nil {
			continue
		}

		typeName := *target.GetOdataType()
		if len(typeName) > len("#microsoft.graph.") && typeName[:len("#microsoft.graph.")] == "#microsoft.graph." {
			typeName = typeName[len("#microsoft.graph."):]
		}

		groupID := types.StringNull()
		if gt, ok := target.(graphmodels.GroupAssignmentTargetable); ok && gt.GetGroupId() != nil {
			groupID = types.StringValue(*gt.GetGroupId())
		}

		obj, diags := types.ObjectValue(AttrTypes, map[string]attr.Value{
			"type":     types.StringValue(typeName),
			"group_id": groupID,
		})
		if diags.HasError() {
			return types.SetNull(objType), fmt.Errorf("building assignment object: %v", diags)
		}
		values = append(values, obj)
	}

	set, diags := types.SetValue(objType, values)
	if diags.HasError() {
		return types.SetNull(objType), fmt.Errorf("building assignments set: %v", diags)
	}
	return set, nil
}

// Assign replaces the full assignment list on a managed app policy.
//
//	POST /deviceAppManagement/managedAppPolicies/{id}/assign
//
// An empty slice removes all assignments.
func Assign(ctx context.Context, client *msgraphbetasdk.GraphServiceClient, policyID string, assignments []graphmodels.TargetedManagedAppPolicyAssignmentable) error {
	adapter := client.GetAdapter()

	body := deviceappmanagement.NewTargetedManagedAppConfigurationsItemAssignPostRequestBody()
	body.SetAssignments(assignments)

	reqInfo := abstractions.NewRequestInformationWithMethodAndUrlTemplateAndPathParameters(
		abstractions.POST,
		assignURLTemplate,
		map[string]string{
			"baseurl":            adapter.GetBaseUrl(),
			"managedAppPolicyId": policyID,
		},
	)
	reqInfo.Headers.TryAdd("Accept", "application/json")
	if err := reqInfo.SetContentFromParsable(ctx, adapter, "application/json", body); err != nil {
		return fmt.Errorf("setting assign request body: %w", err)
	}

	return adapter.SendNoContent(ctx, reqInfo, abstractions.ErrorMappings{
		"XXX": odataerrors.CreateODataErrorFromDiscriminatorValue,
	})
}
