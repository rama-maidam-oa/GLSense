// DataServiceLocator.cs in GLSense.Addin.Core
// Port of GLSense\Service\ServiceLocator.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers).
//
// NAMING NOTE: the old project's own thin data-service layer (Service\ServiceLocator.cs,
// Service\PeriodDataService.cs, Service\SegmentDataService.cs) is a completely different,
// unrelated thing from GLSense.Addin.Core.Infrastructure.ServiceLocator (the
// context/logger/ribbon/Excel wrapper already established by Groups A/B). To avoid two
// types both named "ServiceLocator" colliding by usage confusion within this project, the
// old Service.ServiceLocator is renamed here to DataServiceLocator and moved under the
// Services\ folder/namespace. All internal usages (PeriodDataService/SegmentDataService)
// and the 7 Group C ViewModels' call sites are updated accordingly
// (was: GLSense.Service.ServiceLocator.PeriodDataService.GetPeriodsForLedger(...),
//  now: GLSense.Addin.Core.Services.DataServiceLocator.PeriodDataService.GetPeriodsForLedger(...)).
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;

namespace GLSense.Addin.Core.Services
{
    internal static class DataServiceLocator
    {
        private static IPeriodDataService _periodDataService;
        private static ISegmentDataService _segmentDataService;
        private static readonly object _lock = new object();

        internal static IPeriodDataService PeriodDataService
        {
            get
            {
                lock (_lock)
                {
                    if (_periodDataService == null)
                    {
                        ServiceLocator.Logger?.LogDebug("DataServiceLocator.PeriodDataService: creating new PeriodDataService instance.");
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
                        ServiceLocator.Logger?.LogDebug("DataServiceLocator.SegmentDataService: creating new SegmentDataService instance.");
                        _segmentDataService = new SegmentDataService();
                    }
                    return _segmentDataService;
                }
            }
        }

        // For testing
        internal static void SetPeriodDataService(IPeriodDataService service)
        {
            ServiceLocator.Logger?.LogDebug($"DataServiceLocator.SetPeriodDataService: overriding PeriodDataService instance with {(service == null ? "null" : service.GetType().Name)}.");
            _periodDataService = service;
        }

        internal static void SetSegmentDataService(ISegmentDataService service)
        {
            ServiceLocator.Logger?.LogDebug($"DataServiceLocator.SetSegmentDataService: overriding SegmentDataService instance with {(service == null ? "null" : service.GetType().Name)}.");
            _segmentDataService = service;
        }
    }
}
