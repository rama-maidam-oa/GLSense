// GlobalsEx.cs in GLSense
using GLSense.Contracts;
using GLSense.Loader.Core;

namespace GLSense
{
    public static class GlobalsEx
    {
        public static IGLSenseAddin Addin { get; set; }
        public static AddinDomainLoader Loader { get; set; }
        public static IGLSenseContext Context { get; set; }
    }
}