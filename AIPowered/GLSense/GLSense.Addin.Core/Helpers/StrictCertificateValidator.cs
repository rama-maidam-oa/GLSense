// StrictCertificateValidator.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\StrictCertificateValidator.cs (FinalWorkingCode).
// Changes: LogUtility.* -> ServiceLocator.Logger.*.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Net.Http;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace GLSense.Addin.Core.Helpers
{
    public static class StrictCertificateValidator
    {
        public static bool Validate(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            try
            {
                var cert2 = certificate as X509Certificate2
                            ?? new X509Certificate2(certificate);

                LogCertificate(cert2, chain, sslPolicyErrors);

                // Absolute rule: ANY SSL policy error = FAIL, with the same one
                // exception as the stricter chain check below - if the *only* reason
                // the framework's own automatic chain build flagged an error is that
                // revocation status couldn't be determined, that's not evidence the
                // certificate itself is bad. Always logged at Error (not gated behind
                // Debug mode) with the specific chain status, not just the coarse
                // SslPolicyErrors flag, so the actual cause (untrusted root, partial
                // chain, name mismatch, revocation-lookup-incomplete, etc.) is visible
                // in the log without needing to reproduce with Debug logging enabled.
                // Ported from FinalWorkingCode's identical fix.
                if (sslPolicyErrors != SslPolicyErrors.None)
                {
                    var frameworkStatuses = chain?.ChainStatus;

                    if (frameworkStatuses != null && frameworkStatuses.Length > 0)
                    {
                        foreach (var status in frameworkStatuses)
                        {
                            ServiceLocator.Logger?.LogError(
                                $"TLS validation failed ({sslPolicyErrors}): {status.Status} - {status.StatusInformation}");
                        }
                    }
                    else
                    {
                        ServiceLocator.Logger?.LogError($"TLS validation failed: {sslPolicyErrors}");
                    }

                    // Only revocation-incompleteness is forgiven, and only when it's the
                    // sole policy error - a name mismatch or unavailable certificate
                    // alongside it still fails regardless of what the chain says.
                    bool onlyChainErrorFlag = sslPolicyErrors == SslPolicyErrors.RemoteCertificateChainErrors;
                    if (!onlyChainErrorFlag || !IsOnlyRevocationCheckIncomplete(frameworkStatuses))
                        return false;

                    ServiceLocator.Logger?.LogWarn(
                        "TLS validation: the framework's own chain build flagged only a " +
                        "revocation-status-unknown condition - proceeding anyway; the " +
                        "certificate itself is otherwise valid.");
                }

                // Force Windows chain validation with revocation
                using (var strictChain = new X509Chain())
                {
                    strictChain.ChainPolicy = new X509ChainPolicy
                    {
                        RevocationMode = X509RevocationMode.Online,
                        RevocationFlag = X509RevocationFlag.ExcludeRoot,
                        VerificationFlags = X509VerificationFlags.NoFlag,
                        // Was 300s - long enough to make a soft-failed revocation lookup
                        // (see IsOnlyRevocationCheckIncomplete below) look exactly like a
                        // hang. The online check is a best-effort signal now, so it
                        // doesn't need anywhere near that long before giving up on it.
                        UrlRetrievalTimeout = TimeSpan.FromSeconds(15)
                    };

                    strictChain.ChainPolicy.ApplicationPolicy.Add(
                        new Oid("1.3.6.1.5.5.7.3.1")); // Server Authentication

                    if (!strictChain.Build(cert2))
                    {
                        var statuses = strictChain.ChainStatus;

                        foreach (var status in statuses)
                        {
                            ServiceLocator.Logger?.LogError(
                                $"Chain validation error: {status.Status} - {status.StatusInformation}");
                        }

                        // A genuinely bad certificate (revoked, expired, untrusted root,
                        // wrong key usage, etc.) still fails here exactly as before. The
                        // one condition this does NOT hard-fail on is the revocation
                        // check itself being unable to complete (DNS/firewall/timeout
                        // reaching the CA's OCSP/CRL endpoint - not the customer's own
                        // server). A browser wouldn't hard-fail on that either (OCSP
                        // stapling + soft-fail is the default there), and a corporate
                        // firewall that only allow-lists the app server, not arbitrary
                        // CA infrastructure, otherwise makes this validator fail for
                        // reasons that have nothing to do with whether the certificate
                        // is actually trustworthy.
                        if (IsOnlyRevocationCheckIncomplete(statuses))
                        {
                            ServiceLocator.Logger?.LogWarn(
                                "TLS chain validation: revocation status could not be " +
                                "determined (offline/unreachable OCSP or CRL endpoint) - " +
                                "proceeding anyway; the certificate itself is otherwise valid.");
                        }
                        else
                        {
                            return false;
                        }
                    }
                }

                // Enforce strong crypto (defense-in-depth)
                if (!IsStrongCertificate(cert2))
                {
                    ServiceLocator.Logger?.LogError("Certificate rejected: weak signature or key length");
                    return false;
                }

                string url = TryGetRequestTarget(sender);
                ServiceLocator.Logger?.LogDebug("TLS certificate validation successful for " + url);
                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Fatal TLS validation exception: {ex}");
                return false;
            }
        }

        private static bool IsOnlyRevocationCheckIncomplete(X509ChainStatus[] statuses)
        {
            if (statuses == null || statuses.Length == 0)
                return false;

            const X509ChainStatusFlags revocationIncomplete =
                X509ChainStatusFlags.RevocationStatusUnknown | X509ChainStatusFlags.OfflineRevocation;

            foreach (var status in statuses)
            {
                if ((status.Status & revocationIncomplete) == 0)
                    return false;
            }

            return true;
        }

        private static string TryGetRequestTarget(object sender)
        {
            try
            {
                return sender switch
                {
                    HttpRequestMessage httpRequest => httpRequest.RequestUri?.ToString(),
                    HttpWebRequest webRequest => webRequest.RequestUri?.ToString(),
                    _ => $"Unknown (sender type: {sender?.GetType().FullName})",
                };
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"StrictCertificateValidator.TryGetRequestTarget: could not determine request target - {ex.Message}");
                return string.Empty;
            }
        }

        private static bool IsStrongCertificate(X509Certificate2 cert)
        {
            // RSA < 2048 bits -> reject
            if (cert.PublicKey.Key is System.Security.Cryptography.RSA rsa &&
                rsa.KeySize < 2048)
            {
                return false;
            }

            // Reject weak signature algorithms
            string sigAlg = cert.SignatureAlgorithm.FriendlyName;
            if (sigAlg != null &&
                (sigAlg.IndexOf("md5", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 sigAlg.IndexOf("sha1", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return false;
            }

            return true;
        }

        private static void LogCertificate(
            X509Certificate2 cert,
            X509Chain chain,
            SslPolicyErrors errors)
        {
            ServiceLocator.Logger?.LogDebug("---- TLS CERTIFICATE VALIDATION ----");
            ServiceLocator.Logger?.LogDebug($"Subject: {cert.Subject}");
            ServiceLocator.Logger?.LogDebug($"Issuer: {cert.Issuer}");
            ServiceLocator.Logger?.LogDebug($"Thumbprint: {cert.Thumbprint}");
            ServiceLocator.Logger?.LogDebug($"Valid From: {cert.NotBefore:u}");
            ServiceLocator.Logger?.LogDebug($"Valid To: {cert.NotAfter:u}");
            ServiceLocator.Logger?.LogDebug($"SslPolicyErrors: {errors}");

            if (chain?.ChainStatus != null)
            {
                foreach (var status in chain.ChainStatus)
                {
                    ServiceLocator.Logger?.LogDebug(
                        $"Chain status: {status.Status} - {status.StatusInformation}");
                }
            }

            ServiceLocator.Logger?.LogDebug("-----------------------------------");
        }
    }
}
