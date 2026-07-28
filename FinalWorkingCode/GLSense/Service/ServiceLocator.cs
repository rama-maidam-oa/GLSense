using GLSense.Interfaces;
using GLSense.Utilities;

namespace GLSense.Service
{
    // Simple lazy-singleton service registry for this project (NOT the newer
    // GLSense.Addin.Core Infrastructure/ServiceLocator.cs hot-reload boundary helper).
    // Just resolves/caches the two service instances below; swap-in points for tests
    // are provided via SetPeriodDataService/SetSegmentDataService.
    public static class ServiceLocator
    {
        private static IPeriodDataService _periodDataService;
        private static ISegmentDataService _segmentDataService;
        private static readonly object _lock = new();

        internal static IPeriodDataService PeriodDataService
        {
            get
            {
                lock (_lock)
                {
                    if (_periodDataService == null)
                    {
                        LogUtility.LogDebug("ServiceLocator.PeriodDataService: creating new PeriodDataService instance.");
                        _periodDataService = new PeriodDataService();
                    }
                    return _periodDataService;
                }
            }
        }
        internal static ISegmentDataService SegmentDataService
        {
            get
            {
                lock (_lock)
                {
                    if (_segmentDataService == null)
                    {
                        LogUtility.LogDebug("ServiceLocator.SegmentDataService: creating new SegmentDataService instance.");
                        _segmentDataService = new SegmentDataService();
                    }
                    return _segmentDataService;
                }
            }
        }
        // For testing
        internal static void SetPeriodDataService(IPeriodDataService service)
        {
            LogUtility.LogDebug($"ServiceLocator.SetPeriodDataService: overriding PeriodDataService (test hook). service={(service == null ? "null" : service.GetType().Name)}");
            _periodDataService = service;
        }
        internal static void SetSegmentDataService(ISegmentDataService service)
        {
            LogUtility.LogDebug($"ServiceLocator.SetSegmentDataService: overriding SegmentDataService (test hook). service={(service == null ? "null" : service.GetType().Name)}");
            _segmentDataService = service;
        }
    }
}
