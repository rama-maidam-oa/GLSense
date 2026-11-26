// PostLoginUpdateCheck.cs in GLSense.Addin.Core\Utilities
//
// Login -> add-in update check -> broadcast message. Runs after the login window has closed (GLLogin.SuccessCube).
// The host owns the release index and the reload (IRibbonController.FindNewerAddinRelease / FindNewerPartnerAddinRelease /
// ReloadToNewerAddinRelease). GLSense and - when it is installed too - XLEdge are checked at the same time on background
// threads (Excel stays free for the up to TimeoutSeconds check); then, on Excel's thread, one prompt:
//   - one product has an update: "Add-in updates are available (<product> <version>)... reload?" Yes / No.
//   - both: "Update both" (XLEdge reloads, then GLSense) / "Update GLSense only" / "Not now".
//   Any reload ends the session (it logs out), so the broadcast is skipped; otherwise the broadcast follows as before.
// The check never stops the login: any failure - the server, a timeout, a missing/older XLEdge, or an older host
// without these members (MissingMethodException) - is only logged, and the next step follows.
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
                var own = Task.Run(() => Ask(() => host?.FindNewerAddinRelease(loginUrl, loginToken, TimeoutSeconds), "GLSense"));
                var partner = Task.Run(() => Ask(() => host?.FindNewerPartnerAddinRelease(loginUrl, loginToken, TimeoutSeconds), "XLEdge"));
                Task.WaitAll(own, partner); // Ask never throws
                dispatcher.BeginInvoke(new Action(() => AfterCheck(own.Result, partner.Result, broadcastMsg)));
            });
        }

        private static string Ask(Func<string> check, string product)
        {
            try { return check(); }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: {product} update check skipped ({ex.GetType().Name}: {ex.Message}).");
                return null;
            }
        }

        private static void AfterCheck(string glsense, string xledge, string broadcastMsg)
        {
            try
            {
                if (Prompt(glsense, xledge)) return; // reloading: the session ends, no broadcast
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "PostLoginUpdateCheck: update prompt failed - continuing");
            }

            if (!string.IsNullOrWhiteSpace(broadcastMsg))
                CommonFunctions.GLSenseMessage(broadcastMsg, MessageBoxImage.Information);
        }

        /// <summary>Shows the prompt for whatever has an update; true when a reload was started.</summary>
        private static bool Prompt(string glsense, string xledge)
        {
            bool hasOwn = !string.IsNullOrEmpty(glsense), hasPartner = !string.IsNullOrEmpty(xledge);
            if (!hasOwn && !hasPartner) return false;

            bool includePartner;
            if (hasOwn && hasPartner)
            {
                var answer = CommonFunctions.GLSenseMessage(
                    $"Add-in updates are available:{Environment.NewLine}" +
                    $"    GLSense {glsense}{Environment.NewLine}" +
                    $"    XLEdge {xledge}{Environment.NewLine}{Environment.NewLine}" +
                    $"\"Update both\" reloads GLSense and XLEdge. \"Update GLSense only\" reloads GLSense; XLEdge stays on its current version.",
                    MessageBoxImage.Question,
                    MessageBoxButton.YesNoCancel,
                    new[] { "Update both", "Update GLSense only", "Not now" });
                if (answer != MessageBoxResult.Yes && answer != MessageBoxResult.No)
                {
                    ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: user chose not to update (GLSense {glsense}, XLEdge {xledge}).");
                    return false;
                }
                includePartner = answer == MessageBoxResult.Yes;
            }
            else
            {
                string product = hasOwn ? $"GLSense {glsense}" : $"XLEdge {xledge}";
                var answer = CommonFunctions.GLSenseMessage(
                    $"Add-in updates are available ({product}).{Environment.NewLine}Do you wish to reload the add-in?",
                    MessageBoxImage.Question,
                    MessageBoxButton.YesNo);
                if (answer != MessageBoxResult.Yes)
                {
                    ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: user chose not to reload into {product}.");
                    return false;
                }
                includePartner = hasPartner;
            }

            ServiceLocator.Logger?.LogDebug($"PostLoginUpdateCheck: reloading{(hasOwn ? $" GLSense {glsense}" : "")}{(includePartner ? $" XLEdge {xledge}" : "")}.");
            ServiceLocator.RibbonController?.ReloadToNewerAddinRelease(includePartner);
            return true;
        }
    }
}
