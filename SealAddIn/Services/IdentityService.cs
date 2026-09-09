using System;
using System.DirectoryServices.AccountManagement;

namespace SealAddIn.Services
{
    /// <summary>
    /// VSTOアドインはExcelプロセス内でWindowsサインインユーザーと同一コンテキストで動作するため、
    /// Office SSOのようなクロスプロセス認証は不要です。
    /// </summary>
    public static class IdentityService
    {
        public static string GetOwnerId() => Environment.UserName;

        public static string GetDisplayName()
        {
            var overrideName = Properties.Settings.Default.DisplayNameOverride;
            if (!string.IsNullOrWhiteSpace(overrideName))
            {
                return overrideName;
            }

            try
            {
                using (var context = new PrincipalContext(ContextType.Domain))
                using (var user = UserPrincipal.FindByIdentity(context, Environment.UserName))
                {
                    if (user != null && !string.IsNullOrWhiteSpace(user.DisplayName))
                    {
                        return user.DisplayName;
                    }
                }
            }
            catch
            {
                // ドメイン参照不可の環境(ワークグループ等)ではWindowsアカウント名で代替します。
            }

            return Environment.UserName;
        }
    }
}
