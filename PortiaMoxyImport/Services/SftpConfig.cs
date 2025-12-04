public class SftpConfig
{
    public string Host { get; set; } = string.Empty;
    public int Port { get; set; } = 22;
    public string Username { get; set; } = string.Empty;

    // Optional: password (can be used alone or together with key)
    public string Password { get; set; } = string.Empty;

    // Public key auth
    public string PrivateKeyPath { get; set; } = string.Empty;          // e.g. "C:\\keys\\id_rsa" or "/home/user/.ssh/id_rsa"
    public string PrivateKeyPassphrase { get; set; } = string.Empty;    // leave empty if key is not encrypted

    public string RemoteDirectory { get; set; } = string.Empty;
    public string FilePattern { get; set; } = "*";
    public string LocalDirectory { get; set; } = string.Empty;
}