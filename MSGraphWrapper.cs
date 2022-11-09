using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System;

/// <summary>MSALとMicrosoft GraphをラップしてGraphServiceClientを取得するためのクラス</summary>
public class MSGraphWrapper {
    ////////////////////////////////////////////////////////////////////
    /// <summary>クライアントID</summary>
    private string m_client_id;

    ////////////////////////////////////////////////////////////////////
    /// <summary>テナントID</summary>
    public string m_tenant_id;

    ////////////////////////////////////////////////////////////////////
    /// <summary>リダイレクトURL</summary>
    public string m_redirect_url;

    ////////////////////////////////////////////////////////////////////
    ///<summary></summary>
    private static IPublicClientApplication? m_client_app = null;

    ////////////////////////////////////////////////////////////////////
    ///<summary></summary>
    private static GraphServiceClient? m_auth_client = null;

    ////////////////////////////////////////////////////////////////////
    /// <summary>デフォルトコンストラクタ(呼び出し不可)</summary>
    private MSGraphWrapper() {
        m_client_id = "";
        m_tenant_id = "";
        m_redirect_url = "";
    }

    ////////////////////////////////////////////////////////////////////
    /// <summary>コンストラクタ</summary>
    /// <param name="_client_id">クライアントID</param>
    /// <param name="_tenant_id">テナントID</param>
    /// <param name="_tenant_id">テナントID</param>
    public MSGraphWrapper(string _client_id, string _tenant_id, string _redirect_url) {
        m_client_id = _client_id;
        m_tenant_id = _tenant_id;
        m_redirect_url = _redirect_url;
    }

    ////////////////////////////////////////////////////////////////////
    ///<summary>アクセストークン取得関数</summary>
    /// <param name="_scopes">スコープ(アクセス許可)</param>
    /// <returns>アクセストークン</returns>
    private async Task<string> GetAccessToken(string[] _scopes) {
        AuthenticationResult? auth_ret = null;

        if (m_client_app is null) {
            m_client_app = PublicClientApplicationBuilder.
                        Create(m_client_id).
                        WithAuthority(AzureCloudInstance.AzurePublic, m_tenant_id).
                        WithRedirectUri(m_redirect_url).
                        Build();
            if (m_client_app is null) {
                throw new Exception("IPublicClientApplication Build Error");
            }
        }

        var accounts = await m_client_app.GetAccountsAsync();
        var first_account = accounts?.FirstOrDefault();
        try {
            auth_ret = await m_client_app.AcquireTokenSilent(_scopes, first_account).ExecuteAsync();
        } catch (MsalUiRequiredException) {
            auth_ret = await m_client_app.AcquireTokenInteractive(_scopes).
                                        WithAccount(first_account).
                                        WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount).
                                        ExecuteAsync();
        }

        if (auth_ret is null) {
            throw new Exception("Login Error");
        }

        return auth_ret.AccessToken;
    }

    ////////////////////////////////////////////////////////////////////
    ///<summary>GraphServiceClientの取得</summary>
    /// <param name="_scopes">スコープ(アクセス許可)</param>
    /// <returns>GraphServiceClientインスタンス</returns>
    public async Task<GraphServiceClient> GetGraphClientAsync(string[] _scopes) {
        if (m_auth_client is null) {
            m_auth_client = new GraphServiceClient(
                new DelegateAuthenticationProvider(async (requestMessage) => {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await GetAccessToken(_scopes));
                }));
        }
        return await Task.FromResult(m_auth_client);
    }
}
