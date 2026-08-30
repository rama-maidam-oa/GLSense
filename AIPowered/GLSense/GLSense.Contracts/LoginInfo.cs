using System;

namespace GLSense.Contracts
{
    // [Serializable] because this crosses the host<->Addin.Core AppDomain boundary as
    // the return value of IGLSenseAddin.GetLoginInfo().
    [Serializable]
    public class LoginInfo
    {
        public string LoginUrl { get; set; }
        public string LoginToken { get; set; }
        public bool IsLoggedIn { get; set; }
    }
}
