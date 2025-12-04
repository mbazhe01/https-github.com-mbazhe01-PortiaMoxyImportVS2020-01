using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Renci.SshNet;
using Renci.SshNet.Security;

namespace PortiaMoxyImport.Services
{
    internal class NTSFTPDownloader : IDownloadNTFXTrades
    {
        private readonly SftpConfig _config;

        public NTSFTPDownloader(SftpConfig config)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));
            _config = config;
        }

        public Task<string> DownloadFileAsync()
        {
            // SSH.NET is synchronous, so we wrap it to fit the async interface
            return Task.Run(new Func<string>(DownloadFileInternal));
        }

        private string DownloadFileInternal()
        {
            using (var client = CreateSftpClient())
            {
                client.Connect();
                client.ChangeDirectory(_config.RemoteDirectory);

                var files = client
                    .ListDirectory(".")
                    .Where(f => !f.IsDirectory && WildcardMatch(f.Name, _config.FilePattern))
                    .OrderByDescending(f => f.LastWriteTimeUtc)
                    .ToList();

                if (files.Count == 0)
                {
                    throw new FileNotFoundException(
                        "No files found matching pattern '" + _config.FilePattern + "' in directory '" +
                        _config.RemoteDirectory + "'.");
                }

                var latestFile = files[0];

                if (!Directory.Exists(_config.LocalDirectory))
                {
                    Directory.CreateDirectory(_config.LocalDirectory);
                }

                var localPath = Path.Combine(_config.LocalDirectory, latestFile.Name);

                using (var localFile = File.Open(localPath, FileMode.Create, FileAccess.Write))
                {
                    client.DownloadFile(latestFile.FullName, localFile);
                }

                client.Disconnect();

                return localPath;
            }
        }

        private SftpClient CreateSftpClient()
        {
            // If PrivateKeyPath is provided, use key-based auth (optionally plus password)
            if (!string.IsNullOrWhiteSpace(_config.PrivateKeyPath))
            {
                PrivateKeyFile keyFile;

                if (string.IsNullOrEmpty(_config.PrivateKeyPassphrase))
                {
                    keyFile = new PrivateKeyFile(_config.PrivateKeyPath);
                }
                else
                {
                    keyFile = new PrivateKeyFile(_config.PrivateKeyPath, _config.PrivateKeyPassphrase);
                }

                var authMethods = new System.Collections.Generic.List<AuthenticationMethod>();
                authMethods.Add(new PrivateKeyAuthenticationMethod(_config.Username, keyFile));

                if (!string.IsNullOrEmpty(_config.Password))
                {
                    authMethods.Add(new PasswordAuthenticationMethod(_config.Username, _config.Password));
                }

                var connectionInfo = new ConnectionInfo(
                    _config.Host,
                    _config.Port,
                    _config.Username,
                    authMethods.ToArray());

                return new SftpClient(connectionInfo);
            }

            // Fallback: password-only auth
            return new SftpClient(
                _config.Host,
                _config.Port,
                _config.Username,
                _config.Password);
        }

        // Very simple wildcard matcher supporting '*' anywhere in the pattern.
        private static bool WildcardMatch(string text, string pattern)
        {
            if (string.IsNullOrEmpty(pattern) || pattern == "*")
                return true;

            var parts = pattern.Split('*');
            var position = 0;

            for (var i = 0; i < parts.Length; i++)
            {
                var part = parts[i];
                if (string.IsNullOrEmpty(part))
                    continue;

                var index = text.IndexOf(part, position, StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                    return false;

                position = index + part.Length;
            }

            // If pattern does not end with '*', make sure the last piece is at the end
            if (!pattern.EndsWith("*", StringComparison.Ordinal))
            {
                var lastNonEmpty = parts.LastOrDefault(p => !string.IsNullOrEmpty(p));
                if (!string.IsNullOrEmpty(lastNonEmpty) &&
                    !text.EndsWith(lastNonEmpty, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
