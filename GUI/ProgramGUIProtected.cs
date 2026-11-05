using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Diagnostics;
using System.Security.Principal;
using System.Reflection;
using System.Windows.Forms;
using System.Threading;

namespace ProfileMigratorGuiHost
{
    class Program
    {
        private const byte XOR_KEY = 0xA7;

        [STAThread]
        static int Main(string[] args)
        {
            // Проверка прав администратора
            if (!IsAdministrator())
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = Process.GetCurrentProcess().MainModule.FileName;
                    psi.UseShellExecute = true;
                    psi.Verb = "runas";
                    psi.Arguments = string.Join(" ", args);
                    Process proc = Process.Start(psi);
                    if (proc != null)
                    {
                        proc.WaitForExit();
                        return proc.ExitCode;
                    }
                    return 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Для работы с профилями RDS требуются права Администратора:\n" + ex.Message, "Ошибка запуска", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }
            }

            // Создаем скрытую временную директорию
            string tempDir = Path.Combine(Path.GetTempPath(), "PM_" + Guid.NewGuid().ToString("N"));
            // Имя скрипта обязательно должно заканчиваться на .ps1 для корректной обработки параметром -File
            string tempScript = Path.Combine(tempDir, "pm_host.ps1");

            try
            {
                Directory.CreateDirectory(tempDir);

                // 1. Извлекаем зашифрованный payload из ресурсов сборки
                byte[] decryptedBytes = null;
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.dat"))
                {
                    if (stream == null)
                    {
                        MessageBox.Show("Критическая ошибка: Защищенный модуль payload не найден!", "Ошибка запуска", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return 1;
                    }

                    byte[] encryptedData = new byte[stream.Length];
                    stream.Read(encryptedData, 0, encryptedData.Length);

                    // Снятие XOR-маски
                    for (int i = 0; i < encryptedData.Length; i++)
                    {
                        encryptedData[i] = (byte)(encryptedData[i] ^ XOR_KEY);
                    }

                    // Декомпрессия GZip
                    using (MemoryStream ms = new MemoryStream(encryptedData))
                    using (GZipStream gz = new GZipStream(ms, CompressionMode.Decompress))
                    using (MemoryStream outMs = new MemoryStream())
                    {
                        gz.CopyTo(outMs);
                        decryptedBytes = outMs.ToArray();
                    }
                }

                // 2. Записываем файл в строгом UTF-8 с BOM (EF BB BF)
                // чтобы PowerShell прочитал все русские шрифты и интерфейс со 100% идеальной кодировкой
                using (FileStream fs = new FileStream(tempScript, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    // Гарантируем ровно один маркер UTF-8 BOM
                    if (decryptedBytes.Length < 3 || decryptedBytes[0] != 0xEF || decryptedBytes[1] != 0xBB || decryptedBytes[2] != 0xBF)
                    {
                        fs.Write(new byte[] { 0xEF, 0xBB, 0xBF }, 0, 3);
                    }
                    fs.Write(decryptedBytes, 0, decryptedBytes.Length);
                }
                decryptedBytes = null;

                // 3. Запуск PowerShell
                ProcessStartInfo psInfo = new ProcessStartInfo();
                psInfo.FileName = "powershell.exe";
                psInfo.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -File \"" + tempScript + "\"";
                psInfo.UseShellExecute = false;
                // The script runs from a temporary directory; keep configuration beside the EXE.
                psInfo.EnvironmentVariables["PROFILEMIGRATOR_CONFIG_PATH"] = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pm_config.json");
                psInfo.EnvironmentVariables["PROFILEMIGRATOR_HOST_EXE"] = Process.GetCurrentProcess().MainModule.FileName;
                psInfo.CreateNoWindow = true;
                psInfo.RedirectStandardError = true;
                try {
                    psInfo.StandardErrorEncoding = Encoding.GetEncoding(866);
                } catch {
                    psInfo.StandardErrorEncoding = Encoding.Default;
                }

                using (Process psProcess = Process.Start(psInfo))
                {
                    // Защита исходника: через 1.5 секунды после того как PowerShell зачитал файл в память,
                    // файл удаляется с диска прямо во время работы графического окна!
                    ThreadPool.QueueUserWorkItem((state) =>
                    {
                        try
                        {
                            Thread.Sleep(1500);
                            if (File.Exists(tempScript)) File.Delete(tempScript);
                            if (Directory.Exists(tempDir)) Directory.Delete(tempDir);
                        }
                        catch { }
                    });

                    string errors = psProcess.StandardError.ReadToEnd();
                    psProcess.WaitForExit();

                    if (psProcess.ExitCode != 0 && !string.IsNullOrEmpty(errors))
                    {
                        MessageBox.Show("Сообщение системы:\n\n" + errors, "ProfileMigrator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    return psProcess.ExitCode;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка выполнения:\n" + ex.Message, "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
            finally
            {
                // Финальная зачистка при любом сценарии
                try
                {
                    if (File.Exists(tempScript)) File.Delete(tempScript);
                    if (Directory.Exists(tempDir)) Directory.Delete(tempDir);
                }
                catch { }
            }
        }

        private static bool IsAdministrator()
        {
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }
    }
}
