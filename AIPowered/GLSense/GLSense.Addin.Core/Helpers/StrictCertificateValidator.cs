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

                // Absolute rule: ANY SSL policy error = FAIL
                if (sslPolicyErrors != SslPolicyErrors.None)
                {
                    ServiceLocator.Logger?.LogError($"TLS validation failed: {sslPolicyErrors}");
                    return false;
                }

                // Force Windows chain validation with revocation
                using (var strictChain = new X509Chain())
                {
                    strictChain.ChainPolicy = new X509ChainPolicy
                    {
                        RevocationMode = X509RevocationMode.Online,
                        RevocationFlag = X509RevocationFlag.ExcludeRoot,
                        VerificationFlags = X509VerificationFlags.NoFlag,
                        UrlRetrievalTimeout = TimeSpan.FromSeconds(300)
                    };

                    strictChain.ChainPolicy.ApplicationPolicy.Add(
                        new Oid("1.3.6.1.5.5.7.3.1")); // Server Authentication

                    if (!strictChain.Build(cert2))
                    {
                        foreach (var status in strictChain.ChainStatus)
                        {
                            ServiceLocator.Logger?.LogError(
                                $"Chain validation error: {status.Status} - {status.StatusInformation}");
                        }

                        return false;
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
