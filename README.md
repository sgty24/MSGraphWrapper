# MSGraphWrapper
MSALとMicrosoft GraphをラップしてGraphServiceClientを取得するためのクラス

【サンプル】
``` C#
public partial class Form1 : Form {
    private const string client_id = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";
    private const string tenant_id = "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY";
    private static string[] scopes = { "https://graph.microsoft.com/User.Read" };

    private static MSGraphWrapper m_ms_graph_wrap = new MSGraphWrapper(client_id, tenant_id, "http://localhost");

    private async void button1_Click(object sender, EventArgs e) {
        button1.Enabled = false;
        try {
            var graph_client = await m_ms_graph_wrap.GetGraphClientAsync(scopes);
            var user = await graph_client.Me.Request().GetAsync();
            MessageBox.Show($"DisplayName={user.DisplayName}", "ユーザー名", MessageBoxButtons.OK, MessageBoxIcon.Information);
        } finally {
            button1.Enabled = true;
        }
    }
}
```
