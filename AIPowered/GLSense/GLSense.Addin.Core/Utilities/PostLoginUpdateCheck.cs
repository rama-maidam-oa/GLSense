// PostLoginUpdateCheck.cs in GLSense.Addin.Core\Utilities
//
// Login -> add-in update check -> broadcast message. Runs after the login window has closed (GLLogin.SuccessCube).
// The host owns the release index and the reload (IRibbonController.FindNewerAddinRelease / ReloadToNewerAddinRelease);
// this asks it on a background thread (Excel stays free for the up to TimeoutSeconds check), then, on Excel's thread:
//   - a newer release: "Add-in updates are available... reload?" Yes = the host downloads it, logs out and reloads
//     (the broadcast is skipped - the session is ending); No = carry on.
//   - then the broadcast message, if there is one, exactly as before.
// The check never stops the login: any failure - the server, a timeout, or an older host without these members
// (MissingMethodException) - is only logged, and the broadcast step follows.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Utilities
{
    internal static class PostLoginUpdateCheck
    {
        private const int TimeoutSeconds = 10;

        internal static void Run(Dispatcher dispatcher, string loginUrl, string loginToken, string broadcastMsg)
        {
            var host = ServiceLocator.RibbonController;
            _ = Task.Run(() =>
            {
                string newerRelease = null;
                try
                {
                    newerRelease = host?.FindNewerAddinRelease(loginUrl, loginToken, TimeoutSeconds);
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: update check skipped ({ex.GetType().Name}: {ex.Message}).");
                }

                dispatcher.BeginInvoke(new Action(() => AfterCheck(newerRelease, broadcastMsg)));
            });
        }

        private static void AfterCheck(string newerRelease, string broadcastMsg)
        {
            try
            {
                if (!string.IsNullOrEmpty(newerRelease))
                {
                    var answer = CommonFunctions.GLSenseMessage(
                        $"Add-in updates are available (GLSense {newerRelease}).{Environment.NewLine}Do you wish to reload the add-in?",
                        MessageBoxImage.Question,
                        MessageBoxButton.YesNo);

                    if (answer == MessageBoxResult.Yes)
                    {
                        ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: user chose to reload into {newerRelease}.");
                        ServiceLocator.RibbonController?.ReloadToNewerAddinRelease();
                        return;
                    }

                    ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: user chose not to reload into {newerRelease}.");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PostLoginUpdateCheck: update prompt failed - continuing");
            }

            if (!string.IsNullOrWhiteSpace(broadcastMsg))
                CommonFunctions.GLSenseMessage(broadcastMsg, MessageBoxImage.Information);
        }
    }
}
