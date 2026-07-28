// CompressionHelper.cs in GLSense.Addin.Core
// Ported from GLSense\Helpers\CompressionHelper.cs (FinalWorkingCode) verbatim -
// pure string GZip/base64 compression utility, no Excel/ADX/host dependency.
// Group F (transitive dependency of DataTableBuilder.CreateFormulaKey, which needs it
// to build the compressed formulaKey column cached in AppState.Instance.CalculatedBalances).
// Re-pointed vs. the original: LogUtility.LogException -> ServiceLocator.Logger?.LogException.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace GLSense.Addin.Core.Helpers
{
    public static class CompressionHelper
    {
        private const int MIN_COMPRESSION_SIZE = 10;
        private const string GZIP_HEADER = "H4sI";

        public static string CompressString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            try
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);

                // Skip compression for very small strings
                if (inputBytes.Length < MIN_COMPRESSION_SIZE)
                    return GZIP_HEADER + Convert.ToBase64String(inputBytes);

                using (var outputStream = new MemoryStream())
                {
                    using (var gzipStream = new GZipStream(outputStream, CompressionMode.Compress, true))
                    {
                        gzipStream.Write(inputBytes, 0, inputBytes.Length);
                    }

                    byte[] compressedBytes = outputStream.ToArray();

                    // Check if compression actually reduced size
                    if (compressedBytes.Length >= inputBytes.Length)
                    {
                        // Return uncompressed with header
                        return GZIP_HEADER + Convert.ToBase64String(inputBytes);
                    }

                    // Return compressed with header
                    return GZIP_HEADER + Convert.ToBase64String(compressedBytes);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "CompressionHelper.CompressString");
                // Fallback: return original with header
                return GZIP_HEADER + Convert.ToBase64String(Encoding.UTF8.GetBytes(input));
            }
        }

        public static string DecompressString(string compressedInput)
        {
            if (string.IsNullOrEmpty(compressedInput))
                return compressedInput;

            try
            {
                // If no header, treat as plain text
                if (!compressedInput.StartsWith(GZIP_HEADER, StringComparison.Ordinal))
                    return compressedInput;

                // Strip header first, then decode base64
                string base64Part = compressedInput.Substring(GZIP_HEADER.Length);
                byte[] dataBytes = Convert.FromBase64String(base64Part);

                // Try to decompress; if it is not actually GZip, fall back to UTF-8 text
                using (var inputStream = new MemoryStream(dataBytes))
                {
                    try
                    {
                        using (var gzipStream = new GZipStream(inputStream, CompressionMode.Decompress))
                        using (var outputStream = new MemoryStream())
                        {
                            gzipStream.CopyTo(outputStream);
                            byte[] outputBytes = outputStream.ToArray();
                            if (outputBytes.Length == 0)
                            {
                                // Nothing came out: treat as plain UTF-8
                                return Encoding.UTF8.GetString(dataBytes);
                            }
                            return Encoding.UTF8.GetString(outputBytes);
                        }
                    }
                    catch (InvalidDataException ex)
                    {
                        // Not actually GZip-compressed, just return as UTF-8 string
                        ServiceLocator.Logger?.LogDebug($"CompressionHelper.DecompressString: data was not GZip-compressed, falling back to raw UTF-8 - {ex.Message}");
                        return Encoding.UTF8.GetString(dataBytes);
                    }
                }
            }
            catch (FormatException ex)
            {
                // Not a valid base64 string, return original
                ServiceLocator.Logger?.LogWarn($"CompressionHelper.DecompressString: input was not valid base64, returning input unchanged - {ex.Message}");
                return compressedInput;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "String decompression failed");
                return compressedInput;
            }
        }
    }
}
