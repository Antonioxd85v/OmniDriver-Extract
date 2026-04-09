using System;
using System.Diagnostics;
using System.IO;
using Path = System.IO.Path;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text;
using System.ComponentModel;
using System.Collections.Generic;
using System.Management;
using System.Security.Principal;
using System.Globalization;

// Bibliotecas iText7 para PDF
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Font;
using iText.IO.Font.Constants;

namespace OmniDriver_Extract
{
    public partial class Form1 : Form
    {
        private string idiomaAtual = "ES";
        private CancellationTokenSource _cancellationTokenSource;

        // Acedido apenas na thread de UI (seguro no modelo WinForms)
        private bool _isProcessando = false;

        private enum ChaveMsg
        {
            Analisando, Extraindo, CopiandoGpu, CopiandoImp, GerandoPdf, UacDenied, DismEmpty, Cancelled
        }

        private enum StatusExtracao
        {
            Success, UacDenied, DismEmpty, Cancelled
        }

        private class ResultadoExtracao
        {
            public StatusExtracao Status { get; set; } = StatusExtracao.Success;
            public int Exportados { get; set; } = 0;
            public int SemInf { get; set; } = 0;
            public int Erros { get; set; } = 0;
            public int GpuPacks { get; set; } = 0;
            public string TamanhoFormatado { get; set; } = "0 MB";
            public string StatusImpressora { get; set; } = "N/A";
            public string StatusPdf { get; set; } = "OK";
            public string StatusLog { get; set; } = "OK";
            public string HardwareInfo { get; set; } = "";
        }

        public Form1()
        {
            InitializeComponent();
            this.Text = "OmniDriver Extract - v1.0.0 ENTERPRISE";
            cmbLanguage.Items.AddRange(new string[] { "Español", "Português", "English" });
            cmbLanguage.SelectedIndex = 0;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Fallback de elevação se o app.manifest não estiver configurado
            if (!IsRunningAsAdministrator())
            {
                this.BeginInvoke(new Action(() =>
                {
                    ProcessStartInfo proc = new ProcessStartInfo
                    {
                        UseShellExecute = true,
                        WorkingDirectory = Environment.CurrentDirectory,
                        FileName = Application.ExecutablePath,
                        Verb = "runas"
                    };
                    try { Process.Start(proc); } catch { }
                    Environment.Exit(0);
                }));
                return;
            }
        }

        private bool IsRunningAsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void cmbLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbLanguage.SelectedItem == null) return;
            string sel = cmbLanguage.SelectedItem.ToString();
            idiomaAtual = (sel == "Português") ? "PT" : (sel == "English") ? "EN" : "ES";
            if (!_isProcessando) AtualizarTextosUI();
        }

        private void AtualizarTextosUI()
        {
            lblLanguage.Text = idiomaAtual == "ES" ? "Idioma:" : idiomaAtual == "PT" ? "Idioma:" : "Language:";
            lblPath.Text = idiomaAtual == "ES" ? "Ubicación de guardado:" : idiomaAtual == "PT" ? "Local de gravação:" : "Save location:";
            btnBrowse.Text = idiomaAtual == "ES" ? "Examinar..." : idiomaAtual == "PT" ? "Procurar..." : "Browse...";
            btnExtrair.Text = idiomaAtual == "ES" ? "Exportar Controladores (PDF)" : idiomaAtual == "PT" ? "Exportar Drivers (PDF)" : "Export Drivers (PDF)";
            lblStatus.Text = idiomaAtual == "ES" ? "Listo." : idiomaAtual == "PT" ? "Pronto." : "Ready.";
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private async void btnExtrair_Click(object sender, EventArgs e)
        {
            if (_isProcessando)
            {
                btnExtrair.Enabled = false;
                string txtCancel = idiomaAtual == "PT" ? "A cancelar..." : idiomaAtual == "EN" ? "Canceling..." : "Cancelando...";
                btnExtrair.Text = txtCancel;
                lblStatus.Text = txtCancel;
                _cancellationTokenSource?.Cancel();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtPath.Text))
            {
                MessageBox.Show("Selecciona una carpeta / Seleciona uma pasta", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string idiomaSnapshot = idiomaAtual;
            string destinoRaizSnapshot = txtPath.Text;

            string msgModo = idiomaSnapshot == "PT" ? "Modo Completo: Incluir backup offline da Placa Gráfica? (Requer ~5GB livres)\n\n[Sim] = Modo Completo\n[Não] = Modo Lite (Apenas drivers base)" :
                             idiomaSnapshot == "EN" ? "Complete Mode: Include offline Graphic Card backup? (Requires ~5GB free)\n\n[Yes] = Complete Mode\n[No] = Lite Mode (Base drivers only)" :
                                                      "Modo Completo: ¿Incluir copia offline de la Tarjeta Gráfica? (Requiere ~5GB libres)\n\n[Sí] = Modo Completo\n[No] = Modo Lite (Solo base)";

            DialogResult resultadoModo = MessageBox.Show(msgModo, "OmniDriver - GPU", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (resultadoModo == DialogResult.Cancel) return;
            bool modoCompletoGPUSnapshot = (resultadoModo == DialogResult.Yes);

            string msgPrinter = idiomaSnapshot == "PT" ? "Deseja incluir o backup das configurações de Impressoras?" :
                                idiomaSnapshot == "EN" ? "Do you want to include Printer configurations backup?" :
                                                         "¿Desea incluir la copia de seguridad de la configuración de Impresoras?";

            bool incluirImpressorasSnapshot = (MessageBox.Show(msgPrinter, "OmniDriver - Impressoras", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes);

            try
            {
                DriveInfo drive = new DriveInfo(Path.GetPathRoot(destinoRaizSnapshot));
                long espacoNecessario = modoCompletoGPUSnapshot ? 5368709120L : 1073741824L;
                if (drive.AvailableFreeSpace < espacoNecessario)
                {
                    MessageBox.Show(idiomaSnapshot == "PT" ? "Espaço em disco insuficiente." : "Insufficient disk space.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch { /* Tolerância para UNC/Rede */ }

            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = _cancellationTokenSource.Token;

            PrepararUI(true, idiomaSnapshot);

            var progresso = new Progress<(string Status, int Valor, bool Marquee)>(p =>
            {
                lblStatus.Text = p.Status;
                if (p.Marquee && progressBar1.Style != ProgressBarStyle.Marquee)
                    progressBar1.Style = ProgressBarStyle.Marquee;
                else if (!p.Marquee)
                {
                    if (progressBar1.Style != ProgressBarStyle.Continuous) progressBar1.Style = ProgressBarStyle.Continuous;
                    progressBar1.Value = p.Valor;
                }
            });

            string dataFormatada = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string pdfNome = $"OmniDriver_{Environment.MachineName}_{dataFormatada}.pdf";
            string logNome = $"Export_Log_{Environment.MachineName}_{dataFormatada}.txt";
            string caminhoFinalPDF = Path.Combine(destinoRaizSnapshot, pdfNome);
            string caminhoFinalLog = Path.Combine(destinoRaizSnapshot, logNome);

            try
            {
                ResultadoExtracao resultado = await ExecutarExtracaoAsync(
                    destinoRaizSnapshot,
                    modoCompletoGPUSnapshot,
                    incluirImpressorasSnapshot,
                    idiomaSnapshot,
                    caminhoFinalPDF,
                    caminhoFinalLog,
                    progresso,
                    token);

                if (resultado.Status == StatusExtracao.UacDenied)
                {
                    MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.UacDenied), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (resultado.Status == StatusExtracao.DismEmpty)
                {
                    MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.DismEmpty), "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (resultado.Status == StatusExtracao.Cancelled)
                {
                    MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.Cancelled), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (resultado.Status == StatusExtracao.Success)
                {
                    string resumoMsg = $"✅ Drivers exportados: {resultado.Exportados}\r\n" +
                                       $"⚠️ Sem INF (ignorados): {resultado.SemInf}\r\n" +
                                       $"❌ Erros: {resultado.Erros}\r\n" +
                                       $"📦 Pacotes GPU offline: {resultado.GpuPacks}\r\n" +
                                       $"🖨️ Backup Impressoras: {resultado.StatusImpressora}\r\n" +
                                       $"📄 Relatório PDF: {resultado.StatusPdf}\r\n" +
                                       $"📝 Relatório Log: {resultado.StatusLog}\r\n" +
                                       $"💾 Espaço usado: {resultado.TamanhoFormatado}";

                    MessageBox.Show(resumoMsg, "OmniDriver - Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    try { Process.Start(new ProcessStartInfo(caminhoFinalPDF) { UseShellExecute = true }); } catch { }
                    try { Process.Start("explorer.exe", destinoRaizSnapshot); } catch { }
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.Cancelled), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                PrepararUI(false, idiomaSnapshot);
            }
        }

        // PONTO 1 e 2: Uso de async/await e Tasks que devolvem os valores em vez de fechar variáveis externas
        private async Task<string> ObterInformacaoHardwareAsync(string idioma, CancellationToken mainToken)
        {
            StringBuilder sb = new StringBuilder();
            int pad = 13;

            async Task AddWmiInfoAsync(string label, string query, Func<ManagementBaseObject, string> extractInfo)
            {
                mainToken.ThrowIfCancellationRequested();
                string resultStr = "(N/A)";

                using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(mainToken))
                {
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(5));
                    try
                    {
                        resultStr = await Task.Run(() =>
                        {
                            using (var searcher = new ManagementObjectSearcher(query))
                            using (var col = searcher.Get())
                            {
                                foreach (ManagementBaseObject obj in col)
                                {
                                    using (obj) { return extractInfo(obj); }
                                }
                            }
                            return "(N/A)";
                        }, timeoutCts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        if (mainToken.IsCancellationRequested) throw; // Aborto real pelo utilizador
                        resultStr = "(Timeout)"; // Apenas o WMI bloqueou
                    }
                    catch { resultStr = "(Erro)"; }
                }

                if (string.IsNullOrWhiteSpace(resultStr)) resultStr = "(N/A)";
                sb.AppendLine($"{label.PadRight(pad)}{resultStr}");
            }

            try
            {
                await AddWmiInfoAsync("Modelo:", "SELECT Model FROM Win32_ComputerSystem", obj => obj["Model"]?.ToString());
                await AddWmiInfoAsync("S/N:", "SELECT SerialNumber FROM Win32_BIOS", obj => obj["SerialNumber"]?.ToString());
                await AddWmiInfoAsync("OS:", "SELECT Caption, OSArchitecture FROM Win32_OperatingSystem", obj => $"{obj["Caption"]} ({obj["OSArchitecture"]})");
                await AddWmiInfoAsync("Motherboard:", "SELECT Manufacturer, Product FROM Win32_BaseBoard", obj => $"{obj["Manufacturer"]} {obj["Product"]}");

                // PONTO 3: Extração robusta para sistemas Multi-CPU
                mainToken.ThrowIfCancellationRequested();
                List<string> cpus = new List<string>();
                using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(mainToken))
                {
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(5));
                    try
                    {
                        cpus = await Task.Run(() =>
                        {
                            var list = new List<string>();
                            using (var searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor"))
                            using (var col = searcher.Get())
                            {
                                foreach (ManagementBaseObject obj in col)
                                {
                                    using (obj) { list.Add(obj["Name"]?.ToString() ?? "(N/A)"); }
                                }
                            }
                            return list;
                        }, timeoutCts.Token);
                    }
                    catch (OperationCanceledException) { if (mainToken.IsCancellationRequested) throw; cpus.Add("(Timeout)"); }
                    catch { cpus.Add("(Erro)"); }
                }
                if (cpus.Count == 0) cpus.Add("(N/A)");
                for (int i = 0; i < cpus.Count; i++) { sb.AppendLine($"{(i == 0 ? "CPU:" : "").PadRight(pad)}{cpus[i]}"); }

                await AddWmiInfoAsync("RAM:", "SELECT TotalPhysicalMemory FROM Win32_ComputerSystem", obj => {
                    long ramBytes = Convert.ToInt64(obj["TotalPhysicalMemory"]);
                    int ramGB = (int)Math.Round(ramBytes / 1073741824.0);
                    return $"{ramGB} GB";
                });

                // Tratamento idêntico para a(s) Placa(s) Gráfica(s)
                mainToken.ThrowIfCancellationRequested();
                List<string> gpus = new List<string>();
                using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(mainToken))
                {
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(5));
                    try
                    {
                        gpus = await Task.Run(() =>
                        {
                            var list = new List<string>();
                            using (var searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_VideoController"))
                            using (var col = searcher.Get())
                            {
                                foreach (ManagementBaseObject obj in col)
                                {
                                    using (obj) { list.Add(obj["Name"]?.ToString() ?? "(N/A)"); }
                                }
                            }
                            return list;
                        }, timeoutCts.Token);
                    }
                    catch (OperationCanceledException) { if (mainToken.IsCancellationRequested) throw; gpus.Add("(Timeout)"); }
                    catch { gpus.Add("(Erro)"); }
                }
                if (gpus.Count == 0) gpus.Add("(N/A)");
                for (int i = 0; i < gpus.Count; i++) { sb.AppendLine($"{(i == 0 ? "GPU:" : "").PadRight(pad)}{gpus[i]}"); }
            }
            catch (OperationCanceledException) { throw; }
            catch
            {
                sb.AppendLine(idioma == "PT" ? "Aviso: Falha catastrófica ao ler especificações." :
                              idioma == "EN" ? "Warning: Catastrophic failure reading specifications." :
                                               "Aviso: Fallo catastrófico al leer especificaciones.");
            }
            return sb.ToString();
        }

        // Alterado de sincrono para Async Task devolvendo ResultadoExtracao
        private async Task<ResultadoExtracao> ExecutarExtracaoAsync(
            string destinoRaiz, bool modoGPU, bool modoImp, string idioma,
            string caminhoPDF, string caminhoLog, IProgress<(string, int, bool)> progresso, CancellationToken token)
        {
            ResultadoExtracao result = new ResultadoExtracao();
            result.StatusImpressora = idioma == "PT" ? "Não Solicitado" : idioma == "EN" ? "Not Requested" : "No Solicitado";
            long espacoUsadoBytes = 0;

            Dictionary<string, List<string>> pastasPorCategoria = new Dictionary<string, List<string>>();
            List<(string Pasta, string Cat, string Status, bool Erro)> dadosParaPDF = new List<(string, string, string, bool)>();

            StringBuilder logContent = new StringBuilder();
            logContent.AppendLine($"OmniDriver Export Log - {Environment.MachineName} ({DateTime.Now})");
            logContent.AppendLine("==============================================");

            progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.Analisando), 0, true));

            try
            {
                result.HardwareInfo = await ObterInformacaoHardwareAsync(idioma, token);
            }
            catch (OperationCanceledException)
            {
                result.Status = StatusExtracao.Cancelled;
                return result;
            }

            logContent.AppendLine(result.HardwareInfo);
            logContent.AppendLine("==============================================\r\n");

            string tempPath = Path.Combine(destinoRaiz, "_temp_raw");
            ApagarDiretorioSeguro(tempPath);
            Directory.CreateDirectory(tempPath);

            // A TAREFA DE EXTRAÇÃO DO DISM PASSA A ESTAR DENTRO DO TASK.RUN AQUI
            await Task.Run(() =>
            {
                ProcessStartInfo psiDism = new ProcessStartInfo("dism.exe", $"/online /export-driver /destination:\"{tempPath}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                try
                {
                    using (Process pDism = Process.Start(psiDism))
                    {
                        pDism.OutputDataReceived += (s, ev) => { };
                        pDism.ErrorDataReceived += (s, ev) => { };
                        pDism.BeginOutputReadLine();
                        pDism.BeginErrorReadLine();

                        try
                        {
                            while (!pDism.WaitForExit(500))
                            {
                                token.ThrowIfCancellationRequested();
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            try { if (!pDism.HasExited) pDism.Kill(); } catch { }
                            throw;
                        }

                        if (!Directory.GetDirectories(tempPath).Any())
                        {
                            result.Status = StatusExtracao.DismEmpty;
                            return; // Saída limpa da Task
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    ApagarDiretorioSeguro(tempPath);
                    result.Status = StatusExtracao.Cancelled;
                    return;
                }
                catch (Exception)
                {
                    ApagarDiretorioSeguro(tempPath);
                    result.Status = StatusExtracao.UacDenied;
                    return;
                }

                token.ThrowIfCancellationRequested();

                if (Directory.Exists(tempPath))
                {
                    var pastasExtraidas = Directory.GetDirectories(tempPath);
                    int total = pastasExtraidas.Length;
                    int atual = 0;

                    foreach (var dir in pastasExtraidas)
                    {
                        token.ThrowIfCancellationRequested();
                        atual++;
                        progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.Extraindo), (atual * 100) / total, false));

                        string infFile = Directory.GetFiles(dir, "*.inf", SearchOption.AllDirectories).FirstOrDefault();

                        if (infFile != null)
                        {
                            string[] linhasInf = File.ReadAllLines(infFile);

                            string hwId = ObterHardwareIdAgressivo(linhasInf);
                            string categoria = IdentificarCategoriaCompleta(linhasInf);
                            string fabricante = ObterFabricante(linhasInf);

                            string nomeSugerido = (fabricante != "Component" && !string.IsNullOrWhiteSpace(fabricante)) ? $"{fabricante}_{categoria}" :
                                                  (!string.IsNullOrWhiteSpace(hwId)) ? hwId : $"Unknown_{categoria}";

                            if (nomeSugerido.Length > 50) nomeSugerido = nomeSugerido.Substring(0, 50);

                            string destinoFinal = Path.Combine(destinoRaiz, nomeSugerido);
                            int contador = 1;
                            while (Directory.Exists(destinoFinal))
                            {
                                destinoFinal = Path.Combine(destinoRaiz, $"{nomeSugerido}_v{contador}");
                                contador++;
                            }

                            string nomePastaExibicao = Path.GetFileName(destinoFinal);

                            try
                            {
                                Directory.Move(dir, destinoFinal);
                                espacoUsadoBytes += ObterTamanhoPasta(destinoFinal);

                                if (!pastasPorCategoria.ContainsKey(categoria)) pastasPorCategoria[categoria] = new List<string>();
                                pastasPorCategoria[categoria].Add(nomePastaExibicao);

                                dadosParaPDF.Add((nomePastaExibicao, categoria, "OK", false));
                                logContent.AppendLine($"[OK] {nomePastaExibicao} ({categoria})");
                                result.Exportados++;
                            }
                            catch
                            {
                                dadosParaPDF.Add((nomePastaExibicao, categoria, "ERROR", true));
                                logContent.AppendLine($"[ERROR] Failed to move {nomePastaExibicao}");
                                result.Erros++;
                            }
                        }
                        else
                        {
                            string dirName = Path.GetFileName(dir);
                            dadosParaPDF.Add((dirName, "Unknown", "NO INF", true));
                            logContent.AppendLine($"[IGNORED - NO INF] {dirName}");
                            result.SemInf++;
                            ApagarDiretorioSeguro(dir);
                        }
                    }
                }

                token.ThrowIfCancellationRequested();

                if (modoGPU)
                {
                    progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.CopiandoGpu), 0, true));

                    List<string> prefixosWMI = new List<string>();
                    try
                    {
                        using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController"))
                        using (var col = searcher.Get())
                        {
                            foreach (ManagementBaseObject obj in col)
                            {
                                using (obj)
                                {
                                    string name = obj["Name"]?.ToString().ToLower() ?? "";
                                    if (name.Contains("nvidia")) prefixosWMI.AddRange(new[] { "nv_", "nvd", "nvg", "nvam" });
                                    if (name.Contains("amd") || name.Contains("radeon")) prefixosWMI.AddRange(new[] { "amd", "ati", "radeon" });
                                    if (name.Contains("intel")) prefixosWMI.AddRange(new[] { "iigd", "igdlh" });
                                }
                            }
                        }
                    }
                    catch { }

                    if (prefixosWMI.Count == 0)
                    {
                        prefixosWMI.AddRange(new[] { "nv_", "nvd", "nvg", "nvam", "iigd", "igdlh", "amd", "ati", "radeon" });
                    }

                    string driverStore = Path.Combine(Environment.SystemDirectory, "DriverStore", "FileRepository");
                    if (Directory.Exists(driverStore))
                    {
                        foreach (var dir in Directory.GetDirectories(driverStore))
                        {
                            token.ThrowIfCancellationRequested();
                            string nomePasta = Path.GetFileName(dir).ToLower();
                            if (prefixosWMI.Any(p => nomePasta.StartsWith(p)))
                            {
                                string destinoGPU = Path.Combine(destinoRaiz, "GPU_DriverStore_" + Path.GetFileName(dir));
                                try
                                {
                                    List<string> lockedFiles = new List<string>();
                                    CopiarPastaCompleta(dir, destinoGPU, token, lockedFiles);
                                    espacoUsadoBytes += ObterTamanhoPasta(destinoGPU);

                                    if (lockedFiles.Count == 0)
                                    {
                                        dadosParaPDF.Add((Path.GetFileName(destinoGPU), "GPU_Offline_Pack", "OK", false));
                                        logContent.AppendLine($"[GPU OK] Copied {Path.GetFileName(dir)}");
                                    }
                                    else
                                    {
                                        string lockedStr = string.Join(", ", lockedFiles.Take(3)) + (lockedFiles.Count > 3 ? "..." : "");
                                        dadosParaPDF.Add((Path.GetFileName(destinoGPU), "GPU_Offline_Pack", $"OK ({lockedFiles.Count} Locked)", false));
                                        logContent.AppendLine($"[GPU PARTIAL] Copied {Path.GetFileName(dir)} ({lockedFiles.Count} locked: {lockedStr})");
                                    }

                                    result.GpuPacks++;
                                }
                                catch (OperationCanceledException)
                                {
                                    ApagarDiretorioSeguro(destinoGPU);
                                    throw;
                                }
                                catch (Exception ex)
                                {
                                    dadosParaPDF.Add((Path.GetFileName(destinoGPU), "GPU_Offline_Pack", "ERROR", true));
                                    logContent.AppendLine($"[GPU ERROR] Failed {Path.GetFileName(dir)}: {ex.Message}");
                                    result.Erros++;
                                }
                            }
                        }

                        if (result.GpuPacks > 0)
                        {
                            string batPathGPU = Path.Combine(destinoRaiz, "0_Instalar_GPU_Offline.bat");
                            StringBuilder sbGPU = new StringBuilder();
                            sbGPU.AppendLine("@echo off\r\ntitle OmniDriver - Instalador Offline GPU");
                            sbGPU.AppendLine("net session >nul 2>&1 || (echo ERRO: Execute este ficheiro como Administrador! & pause & exit)");
                            sbGPU.AppendLine("echo ===================================================");
                            sbGPU.AppendLine("echo    Instalando Drivers Completos da Placa Grafica");
                            sbGPU.AppendLine("echo ===================================================");
                            sbGPU.AppendLine("for /D %%G in (\"%~dp0GPU_DriverStore_*\") do (");
                            sbGPU.AppendLine("    echo A instalar: %%~nxG");
                            sbGPU.AppendLine("    pnputil /add-driver \"%%G\\*.inf\" /install /subdirs >nul");
                            sbGPU.AppendLine(")\r\npause >nul");
                            File.WriteAllText(batPathGPU, sbGPU.ToString(), Encoding.Default);
                            espacoUsadoBytes += new FileInfo(batPathGPU).Length;
                        }
                    }
                }

                token.ThrowIfCancellationRequested();

                if (modoImp)
                {
                    progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.CopiandoImp), 0, true));

                    string printerFile = Path.Combine(destinoRaiz, "Backup_Impressoras.printerExport");
                    string printbrmPath = Path.Combine(Environment.SystemDirectory, @"spool\tools\Printbrm.exe");

                    if (File.Exists(printbrmPath))
                    {
                        ProcessStartInfo psiPrint = new ProcessStartInfo(printbrmPath, $"-b -f \"{printerFile}\"")
                        {
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            WindowStyle = ProcessWindowStyle.Hidden
                        };

                        try
                        {
                            using (Process pPrint = Process.Start(psiPrint))
                            {
                                pPrint.OutputDataReceived += (s, ev) => { };
                                pPrint.ErrorDataReceived += (s, ev) => { };
                                pPrint.BeginOutputReadLine();
                                pPrint.BeginErrorReadLine();

                                try
                                {
                                    while (!pPrint.WaitForExit(500))
                                    {
                                        token.ThrowIfCancellationRequested();
                                    }
                                }
                                catch (OperationCanceledException)
                                {
                                    try { if (!pPrint.HasExited) pPrint.Kill(); } catch { }
                                    throw;
                                }
                            }

                            if (File.Exists(printerFile))
                            {
                                dadosParaPDF.Add(("Backup_Impressoras.printerExport", "Printer_Config", "OK", false));
                                logContent.AppendLine($"[PRINTER OK] Configuracoes Exportadas");
                                result.StatusImpressora = "OK (.printerExport)";
                                espacoUsadoBytes += new FileInfo(printerFile).Length;
                            }
                            else
                            {
                                dadosParaPDF.Add(("Backup_Impressoras.printerExport", "Printer_Config", "ERROR", true));
                                logContent.AppendLine($"[PRINTER ERROR] Falha na geracao");
                                result.StatusImpressora = "Erro";
                            }
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            dadosParaPDF.Add(("Backup_Impressoras.printerExport", "Printer_Config", "EXEC ERROR", true));
                            logContent.AppendLine($"[PRINTER ERROR] Falha de execucao do PrintBrm: {ex.Message}");
                            result.StatusImpressora = "Erro Execução";
                        }
                    }
                }

                token.ThrowIfCancellationRequested();
                ApagarDiretorioSeguro(tempPath);

                foreach (var kvp in pastasPorCategoria)
                {
                    string cat = kvp.Key;
                    StringBuilder sbBat = new StringBuilder();
                    sbBat.AppendLine("@echo off\r\ntitle OmniDriver - Instalar " + cat);
                    sbBat.AppendLine("net session >nul 2>&1 || (echo ERRO: Execute Administrador! & pause & exit)");
                    foreach (string pasta in kvp.Value)
                    {
                        sbBat.AppendLine($"pnputil /add-driver \"%~dp0{pasta}\\*.inf\" /install /subdirs >nul");
                    }
                    sbBat.AppendLine("exit");
                    string pathBat = Path.Combine(destinoRaiz, $"Instalar_Categoria_{cat}.bat");
                    File.WriteAllText(pathBat, sbBat.ToString(), Encoding.Default);
                    espacoUsadoBytes += new FileInfo(pathBat).Length;
                }

                if (result.Exportados > 0 || result.GpuPacks > 0)
                {
                    StringBuilder sbMaster = new StringBuilder();
                    sbMaster.AppendLine("@echo off\r\ntitle OmniDriver - Instalacao Completa");
                    sbMaster.AppendLine("net session >nul 2>&1 || (echo ERRO: Execute Administrador! & pause & exit)");
                    sbMaster.AppendLine("for /D %%G in (\"%~dp0*\") do (");
                    sbMaster.AppendLine("    echo %%~nxG | findstr /I /B \"GPU_\" >nul || (");
                    sbMaster.AppendLine("        if exist \"%%G\\*.inf\" ( pnputil /add-driver \"%%G\\*.inf\" /install /subdirs >nul )");
                    sbMaster.AppendLine("    )\r\n)");
                    if (result.GpuPacks > 0)
                    {
                        sbMaster.AppendLine("for /D %%G in (\"%~dp0GPU_DriverStore_*\") do ( pnputil /add-driver \"%%G\\*.inf\" /install /subdirs >nul )");
                    }
                    if (modoImp)
                    {
                        sbMaster.AppendLine("if exist \"%~dp0Backup_Impressoras.printerExport\" (");
                        sbMaster.AppendLine("    \"%SystemRoot%\\System32\\spool\\tools\\PrintBrm.exe\" -r -f \"%~dp0Backup_Impressoras.printerExport\"");
                        sbMaster.AppendLine(")");
                    }
                    sbMaster.AppendLine("pause >nul");
                    string pathMaster = Path.Combine(destinoRaiz, "1_Instalar_TODOS_OS_DRIVERS.bat");
                    File.WriteAllText(pathMaster, sbMaster.ToString(), Encoding.Default);
                    espacoUsadoBytes += new FileInfo(pathMaster).Length;
                }

                progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.GerandoPdf), 0, true));
                try
                {
                    using (PdfWriter writer = new PdfWriter(caminhoPDF))
                    using (PdfDocument pdf = new PdfDocument(writer))
                    using (Document document = new Document(pdf, PageSize.A4))
                    {
                        PdfFont fontBold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                        PdfFont fontMono = PdfFontFactory.CreateFont(StandardFonts.COURIER);

                        document.Add(new Paragraph("OMNIDRIVER - EXPORT REPORT (v1.0.0 ENTERPRISE)")
                            .SetFont(fontBold).SetFontSize(18).SetFontColor(ColorConstants.BLUE)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER));

                        document.Add(new Paragraph($"PC: {Environment.MachineName} | Date: {DateTime.Now} | Mode: {(modoGPU ? "Complete" : "Lite")}\n")
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).SetFontSize(10));

                        string hwTitle = idioma == "PT" ? "Especificações do Sistema:" : idioma == "EN" ? "System Specifications:" : "Especificaciones del Sistema:";
                        document.Add(new Paragraph(hwTitle).SetFont(fontBold).SetFontSize(11).SetFontColor(ColorConstants.DARK_GRAY));
                        document.Add(new Paragraph(result.HardwareInfo).SetFont(fontMono).SetFontSize(9).SetMarginBottom(10));

                        Table table = new Table(UnitValue.CreatePercentArray(new float[] { 50, 30, 20 })).UseAllAvailableWidth();
                        table.AddHeaderCell(new Cell().Add(new Paragraph("Folder Name / HW ID").SetFont(fontBold)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                        table.AddHeaderCell(new Cell().Add(new Paragraph("Category").SetFont(fontBold)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                        table.AddHeaderCell(new Cell().Add(new Paragraph("Status").SetFont(fontBold)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));

                        foreach (var d in dadosParaPDF)
                        {
                            table.AddCell(new Cell().Add(new Paragraph(d.Pasta).SetFontSize(9)));
                            table.AddCell(new Cell().Add(new Paragraph(d.Cat).SetFontSize(8)));

                            var color = d.Erro ? ColorConstants.RED : (d.Status.Contains("NO INF") ? ColorConstants.ORANGE : ColorConstants.GREEN);
                            table.AddCell(new Cell().Add(new Paragraph(d.Status).SetFontSize(8).SetFontColor(color).SetFont(fontBold)));
                        }
                        document.Add(table);
                    }
                    if (File.Exists(caminhoPDF)) espacoUsadoBytes += new FileInfo(caminhoPDF).Length;
                }
                catch (Exception exPdf)
                {
                    result.StatusPdf = "Falhou na Geração";
                    logContent.AppendLine($"[PDF ERROR] {exPdf.Message}");
                }

                result.TamanhoFormatado = (espacoUsadoBytes / (1024.0 * 1024.0)).ToString("0.00") + " MB";
                if (espacoUsadoBytes > 1024 * 1024 * 1024) result.TamanhoFormatado = (espacoUsadoBytes / (1024.0 * 1024.0 * 1024.0)).ToString("0.00") + " GB";

                string finalResumo = $"[RESUMO FINAL]\r\nDrivers Exportados: {result.Exportados}\r\nPacotes GPU: {result.GpuPacks}\r\nSem INF: {result.SemInf}\r\nErros: {result.Erros}\r\nEspaco Usado: {result.TamanhoFormatado}\r\nPDF: {result.StatusPdf}\r\n\r\n";

                try
                {
                    File.WriteAllText(caminhoLog, finalResumo + logContent.ToString(), Encoding.Default);
                }
                catch (Exception exLog)
                {
                    result.StatusLog = $"Falha: {exLog.Message}";
                }

            }, token); // Fim do Task.Run

            return result;
        }

        private void ApagarDiretorioSeguro(string alvo)
        {
            try { if (Directory.Exists(alvo)) Directory.Delete(alvo, true); } catch { }
        }

        private void CopiarPastaCompleta(string origem, string destino, CancellationToken token, List<string> falhas)
        {
            Directory.CreateDirectory(destino);
            foreach (var file in Directory.GetFiles(origem))
            {
                token.ThrowIfCancellationRequested();
                try { File.Copy(file, Path.Combine(destino, Path.GetFileName(file)), true); }
                catch { falhas.Add(Path.GetFileName(file)); }
            }
            foreach (var subDir in Directory.GetDirectories(origem))
            {
                token.ThrowIfCancellationRequested();
                CopiarPastaCompleta(subDir, Path.Combine(destino, Path.GetFileName(subDir)), token, falhas);
            }
        }

        private long ObterTamanhoPasta(string d)
        {
            try { return new DirectoryInfo(d).EnumerateFiles("*.*", SearchOption.AllDirectories).Sum(fi => fi.Length); }
            catch { return 0; }
        }

        private string ObterFabricante(string[] linhasInf)
        {
            try
            {
                var linhaProvider = linhasInf.FirstOrDefault(l => l.TrimStart().StartsWith("Provider", StringComparison.OrdinalIgnoreCase));
                if (linhaProvider != null)
                {
                    string[] partes = linhaProvider.Split(new char[] { '=' }, 2);
                    if (partes.Length > 1)
                    {
                        if (partes[1].Contains("%") || partes[1].Contains("$")) return "Component";

                        string fabricante = partes[1].Replace("\"", "").Trim();
                        foreach (char c in Path.GetInvalidFileNameChars()) { fabricante = fabricante.Replace(c.ToString(), ""); }

                        string fLower = fabricante.ToLower();
                        if (fLower.Contains("amd") || fLower.Contains("advanced micro")) return "AMD";
                        if (fLower.Contains("intel")) return "Intel";
                        if (fLower.Contains("nvidia")) return "NVIDIA";
                        if (fLower.Contains("realtek")) return "Realtek";
                        if (fLower.Contains("microsoft")) return "Microsoft";
                        if (fLower.Contains("synaptics")) return "Synaptics";
                        if (fLower.Contains("logitech")) return "Logitech";
                        if (fLower.Contains("qualcomm")) return "Qualcomm";
                        if (fLower.Contains("mediatek")) return "MediaTek";
                        if (fLower.Contains("broadcom")) return "Broadcom";
                        if (fLower.Contains("marvell")) return "Marvell";
                        if (fLower.Contains("lenovo")) return "Lenovo";
                        if (fLower.Contains("dell")) return "Dell";
                        if (fLower.Contains("hp") || fLower.Contains("hewlett")) return "HP";
                        if (fLower.Contains("epson")) return "Epson";
                        if (fLower.Contains("brother")) return "Brother";
                        if (fLower.Contains("canon")) return "Canon";
                        if (fLower.Contains("lexmark")) return "Lexmark";
                        if (fLower.Contains("xerox")) return "Xerox";

                        if (fabricante.Length > 0 && fabricante.Length < 35) return fabricante;
                    }
                }
            }
            catch { }
            return "Component";
        }

        private string IdentificarCategoriaCompleta(string[] linhasInf)
        {
            try
            {
                string classRaw = null;
                var classLine = linhasInf.FirstOrDefault(l => l.TrimStart().StartsWith("Class=", StringComparison.OrdinalIgnoreCase));

                if (classLine != null)
                {
                    string[] partes = classLine.Split(new char[] { '=' }, 2);
                    if (partes.Length > 1)
                    {
                        classRaw = partes[1].Split(';')[0].Trim().ToLower();
                        foreach (char ch in Path.GetInvalidFileNameChars()) classRaw = classRaw.Replace(ch.ToString(), "");

                        if (classRaw == "net" || classRaw == "nettrans" || classRaw == "netservice") return "Network";
                        if (classRaw == "display") return "Graphics";
                        if (classRaw == "media" || classRaw == "audio") return "Audio";
                        if (classRaw == "system") return "Chipset";
                        if (classRaw == "printer" || classRaw == "printqueue") return "Printer";
                        if (classRaw == "bluetooth") return "Bluetooth";
                        if (classRaw == "camera" || classRaw == "image") return "Camera";
                        if (classRaw == "usb") return "USB";
                        if (classRaw == "mouse") return "Mouse";
                        if (classRaw == "keyboard") return "Keyboard";
                        if (classRaw == "monitor") return "Monitor";
                        if (classRaw == "biometric") return "Biometrics";
                        if (classRaw == "firmware") return "Firmware";
                    }
                }

                var guidLine = linhasInf.FirstOrDefault(l => l.TrimStart().StartsWith("ClassGuid=", StringComparison.OrdinalIgnoreCase));
                if (guidLine != null)
                {
                    string g = guidLine.ToUpperInvariant();
                    if (g.Contains("4D36E972")) return "Network";
                    if (g.Contains("4D36E968")) return "Graphics";
                    if (g.Contains("4D36E96C")) return "Audio";
                    if (g.Contains("4D36E97D")) return "Chipset";
                    if (g.Contains("4D36E979")) return "Printer";
                    if (g.Contains("E0CBF06C")) return "Bluetooth";
                    if (g.Contains("6BDD1FC6")) return "Camera";
                    if (g.Contains("36FC9E60")) return "USB";
                    if (g.Contains("4D36E96F")) return "Mouse";
                    if (g.Contains("4D36E96B")) return "Keyboard";
                    if (g.Contains("4D36E96E")) return "Monitor";
                }

                if (!string.IsNullOrWhiteSpace(classRaw))
                {
                    return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(classRaw);
                }

            }
            catch { }
            return "System";
        }

        private string ObterHardwareIdAgressivo(string[] linhasInf)
        {
            string[] prefixos = { "PCI\\VEN_", "USB\\VID_", "HDAUDIO\\FUNC_", "ACPI\\", "BTHENUM\\", "BTH\\", "HID\\", "DISPLAY\\", "SW\\", "VMBUS\\", "INTELAUDIO\\" };
            try
            {
                foreach (var linha in linhasInf)
                {
                    string lTrim = linha.TrimStart();
                    if (lTrim.StartsWith(";")) continue;
                    if (!lTrim.Contains("\\")) continue;

                    foreach (var p in prefixos)
                    {
                        int index = lTrim.IndexOf(p, StringComparison.OrdinalIgnoreCase);
                        if (index >= 0)
                        {
                            string hwId = lTrim.Substring(index).Split(new char[] { ',', ' ', '\t', '"', '\'' })[0].Trim().ToUpperInvariant();
                            hwId = hwId.Replace("\\", "_").Replace("&", "_");
                            foreach (char c in Path.GetInvalidFileNameChars()) { hwId = hwId.Replace(c.ToString(), ""); }
                            if (hwId.Length > 8) return hwId;
                        }
                    }
                }
            }
            catch { }
            return "";
        }

        private void PrepararUI(bool trabalhando, string idioma)
        {
            _isProcessando = trabalhando;
            btnBrowse.Enabled = !trabalhando;
            cmbLanguage.Enabled = !trabalhando;

            if (trabalhando)
            {
                btnExtrair.Text = idioma == "PT" ? "Cancelar Exportação" : idioma == "EN" ? "Cancel Export" : "Cancelar Exportación";
            }
            else
            {
                AtualizarTextosUI();
                progressBar1.Style = ProgressBarStyle.Blocks;
                progressBar1.Value = 0;
            }
        }

        private string GetMsgTraduzida(string idioma, ChaveMsg chave)
        {
            switch (chave)
            {
                case ChaveMsg.Analisando: return idioma == "PT" ? "A analisar o sistema..." : idioma == "EN" ? "Analyzing system..." : "Analizando sistema...";
                case ChaveMsg.Extraindo: return idioma == "PT" ? "A processar drivers base..." : idioma == "EN" ? "Processing base drivers..." : "Procesando drivers base...";
                case ChaveMsg.CopiandoGpu: return idioma == "PT" ? "A copiar GPU DriverStore..." : idioma == "EN" ? "Copying GPU DriverStore..." : "Copiando GPU DriverStore...";
                case ChaveMsg.CopiandoImp: return idioma == "PT" ? "A guardar config Impressoras..." : idioma == "EN" ? "Saving Printer config..." : "Guardando config Impresoras...";
                case ChaveMsg.GerandoPdf: return idioma == "PT" ? "A gerar relatório PDF..." : idioma == "EN" ? "Generating PDF report..." : "Generando informe PDF...";
                case ChaveMsg.UacDenied: return idioma == "PT" ? "Privilégios de Administrador recusados." : idioma == "EN" ? "Admin privileges denied." : "Permisos de Administrador denegados.";
                case ChaveMsg.DismEmpty: return idioma == "PT" ? "O motor DISM falhou a extração." : idioma == "EN" ? "DISM engine failed extraction." : "El motor DISM falló la extracción.";
                case ChaveMsg.Cancelled: return idioma == "PT" ? "Operação cancelada pelo utilizador." : idioma == "EN" ? "Operation canceled by user." : "Operación cancelada por el usuario.";
                default:
                    System.Diagnostics.Debug.Assert(false, $"ChaveMsg não tratada: {chave}");
                    return chave.ToString();
            }
        }
    }
}