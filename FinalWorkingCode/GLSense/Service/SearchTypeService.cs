using GLSense.Models;
using GLSense.Utilities;
using System.Collections.ObjectModel;

namespace GLSense.Service
{
    // Static, hardcoded lookup lists (no I/O, no failure modes) - logging kept to a
    // minimal entry trace only, same treatment as a lightweight lookup provider.
    public static class SearchTypeService
    {
        private static readonly ObservableCollection<SearchTypeModel> _searchTypes =
            [
                new SearchTypeModel { DisplayName = "Starts With", Value = "StartsWith" },
                new SearchTypeModel { DisplayName = "Does Not Start With", Value = "DoesNotStartWith" },
                new SearchTypeModel { DisplayName = "Ends With", Value = "EndsWith" },
                new SearchTypeModel { DisplayName = "Does Not End With", Value = "DoesNotEndWith" },
                new SearchTypeModel { DisplayName = "Contains", Value = "Contains" },
                new SearchTypeModel { DisplayName = "Not Contains", Value = "NotContains" },
                new SearchTypeModel { DisplayName = "Equals", Value = "Equals" },
                new SearchTypeModel { DisplayName = "Not Equals", Value = "NotEquals" }
            ];

        public static ObservableCollection<SearchTypeModel> GetSearchTypes()
        {
            LogUtility.LogDebug($"SearchTypeService.GetSearchTypes: returning {_searchTypes.Count} search types.");
            return _searchTypes;
        }

        public static SearchTypeModel GetDefaultSearchType()
        {
            LogUtility.LogDebug($"SearchTypeService.GetDefaultSearchType: returning '{_searchTypes[4].DisplayName}'.");
            return _searchTypes[4];
        }
    }
    public static class AttributeTypeService
    {
        private static readonly ObservableCollection<AttributeTypeModel> _attributeTypes =
            [
                new AttributeTypeModel { DisplayName = "ATTRIBUTE1", Value ="ATTRIBUTE1" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE2", Value ="ATTRIBUTE2" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE3", Value ="ATTRIBUTE3" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE4", Value ="ATTRIBUTE4" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE5", Value ="ATTRIBUTE5" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE6", Value ="ATTRIBUTE6" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE7", Value ="ATTRIBUTE7" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE8", Value ="ATTRIBUTE8" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE9", Value ="ATTRIBUTE9" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE10", Value ="ATTRIBUTE10" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE11", Value ="ATTRIBUTE11" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE12", Value ="ATTRIBUTE12" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE13", Value ="ATTRIBUTE13" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE14", Value ="ATTRIBUTE14" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE15", Value ="ATTRIBUTE15" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE16", Value ="ATTRIBUTE16" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE17", Value ="ATTRIBUTE17" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE18", Value ="ATTRIBUTE18" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE19", Value ="ATTRIBUTE19" },
                new AttributeTypeModel { DisplayName = "ATTRIBUTE20", Value ="ATTRIBUTE20" }
            ];

        public static ObservableCollection<AttributeTypeModel> GetAttributesType()
        {
            LogUtility.LogDebug($"AttributeTypeService.GetAttributesType: returning {_attributeTypes.Count} attribute types.");
            return _attributeTypes;
        }

        public static AttributeTypeModel GetDefaultAttributeType()
        {
            LogUtility.LogDebug($"AttributeTypeService.GetDefaultAttributeType: returning '{_attributeTypes[0].DisplayName}'.");
            return _attributeTypes[0];
        }
    }
}
