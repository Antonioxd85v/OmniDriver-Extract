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
        private bool _isProcessando = false;

        private enum ChaveMsg
        {
            Analisando, Extraindo, CopiandoGpu, CopiandoImp, GerandoPdf, UacDenied, DismEmpty, Cancelled, SemDrivers
        }

        private enum ModoOperacao
        {
            Extrair, Restaurar
        }

        // Mensaje interno para logs y depuración; la UI consume GetMsgTraduzida.
        private class DismUacException : Exception
        {
            public DismUacException() : base("UAC_DENIED") { }
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
            public string StatusLog { get; set; } = "N/A"; // N/A por defecto si se cancela antes de llegar al paso de escritura
            public string HardwareInfo { get; set; } = "";
        }

        private class ContextoProcessamento
        {
            public ResultadoExtracao Result { get; set; } = new ResultadoExtracao();
            public string DestinoRaiz { get; set; }
            public string TempPath { get; set; }
            public string Idioma { get; set; }
            public IProgress<(string, int, bool)> Progresso { get; set; }
            public CancellationToken Token { get; set; }
            public StringBuilder LogContent { get; set; } = new StringBuilder();
            public Dictionary<string, List<string>> PastasPorCategoria { get; set; } = new Dictionary<string, List<string>>();
            public List<(string Pasta, string Cat, string Status, bool Erro)> DadosParaPDF { get; set; } = new List<(string, string, string, bool)>();

            private long _espacoUsadoBytes = 0;
            public long EspacoUsadoBytes => Interlocked.Read(ref _espacoUsadoBytes);
            public void AdicionarEspaco(long bytes) => Interlocked.Add(ref _espacoUsadoBytes, bytes);
        }

        public Form1()
        {
            InitializeComponent();
            cmbLanguage.Items.AddRange(new string[] { "Español", "Português", "English" });
            cmbLanguage.SelectedIndex = 0;
            if (!this.Text.Contains("ENTERPRISE")) this.Text = "OmniDriver Extract - v1.0.0 ENTERPRISE";
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

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

            if (btnRestaurar != null)
                btnRestaurar.Text = idiomaAtual == "ES" ? "Restaurar Copia" : idiomaAtual == "PT" ? "Restaurar Backup" : "Restore Backup";
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        // =========================================================================
        // UI: EVENTO DE EXTRACCIÓN
        // =========================================================================
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
                string msgPasta = idiomaAtual == "PT" ? "Selecione uma pasta de destino." :
                                  idiomaAtual == "EN" ? "Select a destination folder." :
                                                        "Seleccione una carpeta de destino.";
                string msgTitulo = idiomaAtual == "EN" ? "Warning" : "Aviso";
                MessageBox.Show(msgPasta, msgTitulo, MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    string msgEspaco = idiomaSnapshot == "PT" ? "Espaço em disco insuficiente." :
                                       idiomaSnapshot == "EN" ? "Insufficient disk space." :
                                                                "Espacio en disco insuficiente.";
                    string msgTituloErro = idiomaSnapshot == "EN" ? "Error" : "Erro";
                    MessageBox.Show(msgEspaco, msgTituloErro, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch { /* Tolerancia para rutas UNC/Red */ }

            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = _cancellationTokenSource.Token;

            PrepararUI(true, idiomaSnapshot, ModoOperacao.Extrair);

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
                ResultadoExtracao resultado = await MotorPrincipalExtracao(
                    destinoRaizSnapshot,
                    modoCompletoGPUSnapshot,
                    incluirImpressorasSnapshot,
                    idiomaSnapshot,
                    caminhoFinalPDF,
                    caminhoFinalLog,
                    progresso,
                    token);

                if (resultado.Exportados == 0 && resultado.GpuPacks == 0)
                {
                    string tituloAviso = idiomaSnapshot == "EN" ? "Warning" : "Aviso";
                    MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.SemDrivers), tituloAviso, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string resumoMsg = idiomaSnapshot == "EN" ?
                        $"✅ Exported Drivers: {resultado.Exportados}\r\n" +
                        $"⚠️ No INF (ignored): {resultado.SemInf}\r\n" +
                        $"❌ Errors: {resultado.Erros}\r\n" +
                        $"📦 Offline GPU Packs: {resultado.GpuPacks}\r\n" +
                        $"🖨️ Printers Backup: {resultado.StatusImpressora}\r\n" +
                        $"📄 PDF Report: {resultado.StatusPdf}\r\n" +
                        $"📝 Log Report: {resultado.StatusLog}\r\n" +
                        $"💾 Space Used: {resultado.TamanhoFormatado}"
                        : idiomaSnapshot == "ES" ?
                        $"✅ Controladores exportados: {resultado.Exportados}\r\n" +
                        $"⚠️ Sin INF (ignorados): {resultado.SemInf}\r\n" +
                        $"❌ Errores: {resultado.Erros}\r\n" +
                        $"📦 Paquetes GPU offline: {resultado.GpuPacks}\r\n" +
                        $"🖨️ Copia Impresoras: {resultado.StatusImpressora}\r\n" +
                        $"📄 Informe PDF: {resultado.StatusPdf}\r\n" +
                        $"📝 Informe Log: {resultado.StatusLog}\r\n" +
                        $"💾 Espacio usado: {resultado.TamanhoFormatado}"
                        : // PT
                        $"✅ Drivers exportados: {resultado.Exportados}\r\n" +
                        $"⚠️ Sem INF (ignorados): {resultado.SemInf}\r\n" +
                        $"❌ Erros: {resultado.Erros}\r\n" +
                        $"📦 Pacotes GPU offline: {resultado.GpuPacks}\r\n" +
                        $"🖨️ Backup Impressoras: {resultado.StatusImpressora}\r\n" +
                        $"📄 Relatório PDF: {resultado.StatusPdf}\r\n" +
                        $"📝 Relatório Log: {resultado.StatusLog}\r\n" +
                        $"💾 Espaço usado: {resultado.TamanhoFormatado}";

                    string tituloSucesso = idiomaSnapshot == "EN" ? "OmniDriver - Success" : idiomaSnapshot == "ES" ? "OmniDriver - Éxito" : "OmniDriver - Sucesso";
                    MessageBox.Show(resumoMsg, tituloSucesso, MessageBoxButtons.OK, MessageBoxIcon.Information);

                    try { Process.Start(new ProcessStartInfo(caminhoFinalPDF) { UseShellExecute = true }); } catch { }
                    try { Process.Start("explorer.exe", destinoRaizSnapshot); } catch { }
                }
            }
            catch (OperationCanceledException)
            {
                string tituloAviso = idiomaSnapshot == "EN" ? "Warning" : "Aviso";
                MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.Cancelled), tituloAviso, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (DismUacException)
            {
                string tituloAviso = idiomaSnapshot == "EN" ? "Warning" : "Aviso";
                MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.UacDenied), tituloAviso, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (InvalidOperationException)
            {
                string tituloErro = idiomaSnapshot == "EN" ? "Error" : "Erro";
                MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.DismEmpty), tituloErro, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                string tituloCritico = idiomaSnapshot == "EN" ? "Critical Error" : idiomaSnapshot == "ES" ? "Error Crítico" : "Erro Crítico";
                MessageBox.Show(ex.Message, tituloCritico, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                PrepararUI(false, idiomaSnapshot, ModoOperacao.Extrair);
            }
        }

        // =========================================================================
        // UI: EVENTO DE RESTAURACIÓN
        // =========================================================================
        private async void btnRestaurar_Click(object sender, EventArgs e)
        {
            if (_isProcessando)
            {
                btnRestaurar.Enabled = false;
                string txtCancel = idiomaAtual == "PT" ? "A cancelar..." : idiomaAtual == "EN" ? "Canceling..." : "Cancelando...";
                btnRestaurar.Text = txtCancel;
                lblStatus.Text = txtCancel;
                _cancellationTokenSource?.Cancel();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtPath.Text))
            {
                string msgPasta = idiomaAtual == "PT" ? "Selecione a pasta do backup." :
                                  idiomaAtual == "EN" ? "Select the backup folder." :
                                                        "Seleccione la carpeta de la copia.";
                string msgTitulo = idiomaAtual == "EN" ? "Warning" : "Aviso";
                MessageBox.Show(msgPasta, msgTitulo, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string idiomaSnapshot = idiomaAtual;
            string pastaBackupSnapshot = txtPath.Text;

            // Validación heurística super rápida
            bool temBatMaster = File.Exists(Path.Combine(pastaBackupSnapshot, "1_Instalar_TODOS_OS_DRIVERS.bat"));
            bool temLogExportacao = Directory.GetFiles(pastaBackupSnapshot, "Export_Log_*.txt").Any();
            bool temSubPastasComInf = Directory.Exists(pastaBackupSnapshot) &&
                                      Directory.GetDirectories(pastaBackupSnapshot)
                                      .Any(d => Directory.GetFiles(d, "*.inf", SearchOption.TopDirectoryOnly).Any());

            if (!temBatMaster && !temLogExportacao && !temSubPastasComInf)
            {
                string msgAvisoEstructura = idiomaSnapshot == "PT" ? "Aviso: Estrutura de backup não reconhecida ou sem drivers válidos." :
                                            idiomaSnapshot == "EN" ? "Warning: Backup structure not recognized or no valid drivers." :
                                                                     "Aviso: Estructura de copia no reconocida o sin controladores válidos.";
                string msgTituloErro = idiomaSnapshot == "EN" ? "Error" : "Erro";
                MessageBox.Show(msgAvisoEstructura, msgTituloErro, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string msgModo = idiomaSnapshot == "PT" ? "Deseja restaurar os pacotes pesados da Placa Gráfica (se existirem)?" :
                             idiomaSnapshot == "EN" ? "Do you want to restore heavy GPU packs (if any)?" :
                                                      "¿Desea restaurar los paquetes pesados de la Tarjeta Gráfica (si existen)?";
            DialogResult resultadoModo = MessageBox.Show(msgModo, "OmniDriver - GPU", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (resultadoModo == DialogResult.Cancel) return;
            bool restaurarGPUSnapshot = (resultadoModo == DialogResult.Yes);

            string msgPrinter = idiomaSnapshot == "PT" ? "Deseja restaurar as configurações de Impressoras?" :
                                idiomaSnapshot == "EN" ? "Do you want to restore Printer configurations?" :
                                                         "¿Desea restaurar la configuración de Impresoras?";
            bool restaurarImpSnapshot = (MessageBox.Show(msgPrinter, "OmniDriver - Impressoras", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes);

            _cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = _cancellationTokenSource.Token;

            PrepararUI(true, idiomaSnapshot, ModoOperacao.Restaurar);

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

            try
            {
                ResultadoExtracao resultado = await MotorPrincipalRestauro(
                    pastaBackupSnapshot,
                    restaurarGPUSnapshot,
                    restaurarImpSnapshot,
                    idiomaSnapshot,
                    progresso,
                    token);

                if (resultado.Exportados == 0 && resultado.GpuPacks == 0 && !restaurarImpSnapshot)
                {
                    string msgSemDrivers = idiomaSnapshot == "PT" ? "Nenhum driver instalado." :
                                           idiomaSnapshot == "EN" ? "No drivers installed." :
                                                                    "Ningún controlador instalado.";
                    string tituloAviso = idiomaSnapshot == "EN" ? "Warning" : "Aviso";
                    MessageBox.Show(msgSemDrivers, tituloAviso, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string footerReboot = idiomaSnapshot == "PT" ? "É altamente recomendado reiniciar o computador." :
                                          idiomaSnapshot == "EN" ? "A system reboot is highly recommended." :
                                                                   "Se recomienda encarecidamente reiniciar el equipo.";

                    string resumoMsg = idiomaSnapshot == "EN" ?
                        $"✅ Folders Installed: {resultado.Exportados}\r\n" +
                        $"❌ Installation Errors: {resultado.Erros}\r\n" +
                        $"📦 GPU Packs Installed: {resultado.GpuPacks}\r\n" +
                        $"🖨️ Printers Restored: {resultado.StatusImpressora}\r\n" +
                        $"📝 Log Report: {resultado.StatusLog}\r\n\r\n{footerReboot}"
                        : idiomaSnapshot == "ES" ?
                        $"✅ Carpetas instaladas: {resultado.Exportados}\r\n" +
                        $"❌ Errores de instalación: {resultado.Erros}\r\n" +
                        $"📦 Paquetes GPU instalados: {resultado.GpuPacks}\r\n" +
                        $"🖨️ Impresoras restauradas: {resultado.StatusImpressora}\r\n" +
                        $"📝 Informe Log: {resultado.StatusLog}\r\n\r\n{footerReboot}"
                        : // PT
                        $"✅ Pastas Instaladas: {resultado.Exportados}\r\n" +
                        $"❌ Erros de Instalação: {resultado.Erros}\r\n" +
                        $"📦 Pacotes GPU Instalados: {resultado.GpuPacks}\r\n" +
                        $"🖨️ Restauro Impressoras: {resultado.StatusImpressora}\r\n" +
                        $"📝 Relatório Log: {resultado.StatusLog}\r\n\r\n{footerReboot}";

                    string tituloSucesso = idiomaSnapshot == "EN" ? "OmniDriver - Restore Complete" : idiomaSnapshot == "ES" ? "OmniDriver - Restauración Completada" : "OmniDriver - Restauro Concluído";
                    MessageBox.Show(resumoMsg, tituloSucesso, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                string tituloAviso = idiomaSnapshot == "EN" ? "Warning" : "Aviso";
                MessageBox.Show(GetMsgTraduzida(idiomaSnapshot, ChaveMsg.Cancelled), tituloAviso, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                string tituloCritico = idiomaSnapshot == "EN" ? "Critical Error" : idiomaSnapshot == "ES" ? "Error Crítico" : "Erro Crítico";
                MessageBox.Show(ex.Message, tituloCritico, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                PrepararUI(false, idiomaSnapshot, ModoOperacao.Restaurar);
            }
        }

        // =========================================================================
        // MOTOR DE RESTAURACIÓN (IMPLEMENTACIÓN)
        // =========================================================================
        private async Task<ResultadoExtracao> MotorPrincipalRestauro(
            string pastaBackup, bool restaurarGPU, bool restaurarImp, string idioma,
            IProgress<(string, int, bool)> progresso, CancellationToken token)
        {
            ResultadoExtracao result = new ResultadoExtracao();
            result.StatusImpressora = idioma == "PT" ? "Não Solicitado" : idioma == "EN" ? "Not Requested" : "No Solicitado";

            // Marca de tiempo inmutable para el log
            string timestampInicio = DateTime.Now.ToString("yyyyMMdd_HHmm");

            StringBuilder logRestauro = new StringBuilder();
            logRestauro.AppendLine($"OmniDriver Restore Log - {Environment.MachineName} ({DateTime.Now})");
            logRestauro.AppendLine("==============================================");

            progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.Analisando), 0, true));

            await Task.Run(() =>
            {
                try
                {
                    token.ThrowIfCancellationRequested();

                    var subDirs = Directory.GetDirectories(pastaBackup);

                    var baseDirs = subDirs.Where(d => !Path.GetFileName(d).StartsWith("GPU_DriverStore_", StringComparison.OrdinalIgnoreCase))
                                          .Where(d => Directory.GetFiles(d, "*.inf", SearchOption.AllDirectories).Any())
                                          .ToList();

                    var gpuDirs = subDirs.Where(d => Path.GetFileName(d).StartsWith("GPU_DriverStore_", StringComparison.OrdinalIgnoreCase))
                                         .Where(d => Directory.GetFiles(d, "*.inf", SearchOption.AllDirectories).Any())
                                         .ToList();

                    if (baseDirs.Count > 0)
                    {
                        logRestauro.AppendLine("\n--- Instalando Drivers Base ---");
                        int instalados = 0;

                        foreach (string dir in baseDirs)
                        {
                            token.ThrowIfCancellationRequested();
                            instalados++;
                            int percentagem = (instalados * 100) / baseDirs.Count;

                            string nomePasta = Path.GetFileName(dir);
                            string msg = GetMsgInstalando(idioma, nomePasta); // Formato Tipado y Seguro

                            progresso.Report((msg, percentagem, false));

                            if (InstalarDriverPnP(dir, token))
                            {
                                result.Exportados++;
                                logRestauro.AppendLine($"[OK] {nomePasta}");
                            }
                            else
                            {
                                result.Erros++;
                                logRestauro.AppendLine($"[ERROR] Falha ao instalar {nomePasta}");
                            }
                        }
                    }

                    if (restaurarGPU && gpuDirs.Count > 0)
                    {
                        logRestauro.AppendLine("\n--- Instalando Pacotes GPU ---");
                        progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.CopiandoGpu), 0, true));

                        foreach (string dirGpu in gpuDirs)
                        {
                            token.ThrowIfCancellationRequested();
                            string nomeGpu = Path.GetFileName(dirGpu);
                            if (InstalarDriverPnP(dirGpu, token))
                            {
                                result.GpuPacks++;
                                logRestauro.AppendLine($"[OK] {nomeGpu}");
                            }
                            else
                            {
                                result.Erros++;
                                logRestauro.AppendLine($"[ERROR] Falha no pacote GPU: {nomeGpu}");
                            }
                        }
                    }

                    if (restaurarImp)
                    {
                        logRestauro.AppendLine("\n--- Restaurando Impressoras ---");
                        token.ThrowIfCancellationRequested();
                        progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.CopiandoImp), 0, true));

                        string printerFile = Path.Combine(pastaBackup, "Backup_Impressoras.printerExport");
                        if (File.Exists(printerFile))
                        {
                            if (RestaurarImpressorasPrintBrm(printerFile, token))
                            {
                                result.StatusImpressora = "OK (Restaurado)";
                                logRestauro.AppendLine("[OK] Configuracoes de impressoras restauradas");
                            }
                            else
                            {
                                result.StatusImpressora = "Erro no Restauro";
                                logRestauro.AppendLine("[ERROR] PrintBrm retornou erro durante o restauro");
                            }
                        }
                        else
                        {
                            result.StatusImpressora = "Ficheiro não encontrado";
                            logRestauro.AppendLine("[IGNORED] Ficheiro Backup_Impressoras.printerExport nao encontrado");
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    result.Status = StatusExtracao.Cancelled;
                    logRestauro.AppendLine("\n>> OPERACAO CANCELADA PELO UTILIZADOR <<");
                    throw;
                }
                catch (Exception ex)
                {
                    logRestauro.AppendLine($"\n>> ERRO FATAL: {ex.Message} <<");
                    throw;
                }
                finally
                {
                    try
                    {
                        string logName = $"Restore_Log_{Environment.MachineName}_{timestampInicio}.txt";
                        File.WriteAllText(Path.Combine(pastaBackup, logName), logRestauro.ToString(), Encoding.Default);
                        result.StatusLog = "OK";
                    }
                    catch (Exception exLog)
                    {
                        result.StatusLog = $"Falha IO: {exLog.Message}";
                    }
                }
            }, token);

            return result;
        }

        private bool InstalarDriverPnP(string caminhoPasta, CancellationToken token)
        {
            ProcessStartInfo psi = new ProcessStartInfo("pnputil.exe", $"/add-driver \"{caminhoPasta}\\*.inf\" /install /subdirs")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            try
            {
                using (Process p = Process.Start(psi))
                {
                    p.OutputDataReceived += (s, ev) => { };
                    p.ErrorDataReceived += (s, ev) => { };
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();

                    try
                    {
                        while (!p.WaitForExit(500))
                        {
                            token.ThrowIfCancellationRequested();
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        try { if (!p.HasExited) p.Kill(); } catch { }
                        throw;
                    }
                    return p.ExitCode == 0 || p.ExitCode == 259;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch { return false; }
        }

        private bool RestaurarImpressorasPrintBrm(string caminhoBackup, CancellationToken token)
        {
            string printbrmPath = Path.Combine(Environment.SystemDirectory, @"spool\tools\Printbrm.exe");
            if (!File.Exists(printbrmPath)) return false;

            ProcessStartInfo psi = new ProcessStartInfo(printbrmPath, $"-r -f \"{caminhoBackup}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            try
            {
                using (Process p = Process.Start(psi))
                {
                    p.OutputDataReceived += (s, ev) => { };
                    p.ErrorDataReceived += (s, ev) => { };
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();

                    try
                    {
                        while (!p.WaitForExit(500))
                        {
                            token.ThrowIfCancellationRequested();
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        try { if (!p.HasExited) p.Kill(); } catch { }
                        throw;
                    }
                    return p.ExitCode == 0;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch { return false; }
        }

        // =========================================================================
        // MOTOR DE EXTRACCIÓN
        // =========================================================================
        private async Task<ResultadoExtracao> MotorPrincipalExtracao(
            string destinoRaiz, bool modoGPU, bool modoImp, string idioma,
            string caminhoPDF, string caminhoLog, IProgress<(string, int, bool)> progresso, CancellationToken token)
        {
            ContextoProcessamento ctx = new ContextoProcessamento
            {
                DestinoRaiz = destinoRaiz,
                TempPath = Path.Combine(destinoRaiz, "_temp_raw"),
                Idioma = idioma,
                Progresso = progresso,
                Token = token
            };

            ctx.Result.StatusImpressora = idioma == "PT" ? "Não Solicitado" : idioma == "EN" ? "Not Requested" : "No Solicitado";
            ctx.LogContent.AppendLine($"OmniDriver Export Log - {Environment.MachineName} ({DateTime.Now})");
            ctx.LogContent.AppendLine("==============================================");

            progresso.Report((GetMsgTraduzida(idioma, ChaveMsg.Analisando), 0, true));
            ctx.Result.HardwareInfo = await ObterInformacaoHardwareAsync(idioma, token);

            ctx.LogContent.AppendLine(ctx.Result.HardwareInfo);
            ctx.LogContent.AppendLine("==============================================\r\n");

            await Task.Run(() =>
            {
                try
                {
                    Passo1_ExtrairDism(ctx);
                    Passo2_OrganizarDrivers(ctx);

                    if (modoGPU) Passo3_BackupGpuOffline(ctx);
                    if (modoImp) Passo4_BackupImpressoras(ctx);

                    Passo5_GerarScriptsBat(ctx, modoGPU, modoImp);
                    Passo6_FinalizarRelatorios(ctx, caminhoPDF, caminhoLog);
                }
                finally
                {
                    ApagarDiretorioSeguro(ctx.TempPath);
                }
            }, token);

            return ctx.Result;
        }

        // =========================================================================
        // WMI HARDWARE EXTRACTION
        // =========================================================================
        private async Task<string> ObterInformacaoHardwareAsync(string idioma, CancellationToken mainToken)
        {
            StringBuilder sb = new StringBuilder();
            int pad = 14;

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
                        if (mainToken.IsCancellationRequested) throw;
                        resultStr = "(Timeout)";
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
                await AddWmiInfoAsync("BIOS:", "SELECT Manufacturer, SMBIOSBIOSVersion FROM Win32_BIOS", obj => $"{obj["Manufacturer"]} {obj["SMBIOSBIOSVersion"]}");
                await AddWmiInfoAsync("OS:", "SELECT Caption, OSArchitecture FROM Win32_OperatingSystem", obj => $"{obj["Caption"]} ({obj["OSArchitecture"]})");
                await AddWmiInfoAsync("Motherboard:", "SELECT Manufacturer, Product FROM Win32_BaseBoard", obj => $"{obj["Manufacturer"]} {obj["Product"]}");

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

                mainToken.ThrowIfCancellationRequested();
                List<string> discos = new List<string>();
                using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(mainToken))
                {
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(5));
                    try
                    {
                        discos = await Task.Run(() =>
                        {
                            var list = new List<string>();
                            using (var searcher = new ManagementObjectSearcher("SELECT Model, Size FROM Win32_DiskDrive"))
                            using (var col = searcher.Get())
                            {
                                foreach (ManagementBaseObject obj in col)
                                {
                                    using (obj)
                                    {
                                        if (obj["Size"] != null)
                                        {
                                            long bytes = Convert.ToInt64(obj["Size"]);
                                            int gb = (int)Math.Round(bytes / 1073741824.0);
                                            string tamanhoFormatado = gb >= 1000 ? $"{Math.Round(gb / 1000.0, 1)} TB" : $"{gb} GB";
                                            list.Add($"{obj["Model"]} ({tamanhoFormatado})");
                                        }
                                        else { list.Add($"{obj["Model"]} (Sem Media)"); }
                                    }
                                }
                            }
                            return list;
                        }, timeoutCts.Token);
                    }
                    catch (OperationCanceledException) { if (mainToken.IsCancellationRequested) throw; discos.Add("(Timeout)"); }
                    catch { discos.Add("(Erro)"); }
                }
                if (discos.Count == 0) discos.Add("(N/A)");
                for (int i = 0; i < discos.Count; i++) { sb.AppendLine($"{(i == 0 ? "Disco(s):" : "").PadRight(pad)}{discos[i]}"); }

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

                await AddWmiInfoAsync("Bateria:", "SELECT EstimatedChargeRemaining FROM Win32_Battery", obj => $"{obj["EstimatedChargeRemaining"]}%");
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

        // ==========================================
        // PASOS DE LA EXTRACCIÓN (IMPLEMENTACIÓN)
        // ==========================================
        private void Passo1_ExtrairDism(ContextoProcessamento ctx)
        {
            ctx.Progresso.Report((GetMsgTraduzida(ctx.Idioma, ChaveMsg.Analisando), 0, true));
            ApagarDiretorioSeguro(ctx.TempPath);
            Directory.CreateDirectory(ctx.TempPath);

            ProcessStartInfo psiDism = new ProcessStartInfo("dism.exe", $"/online /export-driver /destination:\"{ctx.TempPath}\"")
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
                            ctx.Token.ThrowIfCancellationRequested();
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        try { if (!pDism.HasExited) pDism.Kill(); } catch { }
                        throw;
                    }

                    if (!Directory.GetDirectories(ctx.TempPath).Any())
                        throw new InvalidOperationException("DISM_EMPTY");
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (InvalidOperationException) { throw; }
            catch { throw new DismUacException(); }
        }

        private void Passo2_OrganizarDrivers(ContextoProcessamento ctx)
        {
            ctx.Token.ThrowIfCancellationRequested();

            if (!Directory.Exists(ctx.TempPath)) return;

            var pastasExtraidas = Directory.GetDirectories(ctx.TempPath);
            int total = pastasExtraidas.Length;
            int atual = 0;

            foreach (var dir in pastasExtraidas)
            {
                ctx.Token.ThrowIfCancellationRequested();
                atual++;
                ctx.Progresso.Report((GetMsgTraduzida(ctx.Idioma, ChaveMsg.Extraindo), (atual * 100) / total, false));

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

                    string destinoFinal = Path.Combine(ctx.DestinoRaiz, nomeSugerido);
                    int contador = 1;
                    while (Directory.Exists(destinoFinal))
                    {
                        destinoFinal = Path.Combine(ctx.DestinoRaiz, $"{nomeSugerido}_v{contador}");
                        contador++;
                    }

                    string nomePastaExibicao = Path.GetFileName(destinoFinal);

                    try
                    {
                        Directory.Move(dir, destinoFinal);
                        ctx.AdicionarEspaco(ObterTamanhoPasta(destinoFinal));

                        if (!ctx.PastasPorCategoria.ContainsKey(categoria)) ctx.PastasPorCategoria[categoria] = new List<string>();
                        ctx.PastasPorCategoria[categoria].Add(nomePastaExibicao);

                        ctx.DadosParaPDF.Add((nomePastaExibicao, categoria, "OK", false));
                        ctx.LogContent.AppendLine($"[OK] {nomePastaExibicao} ({categoria})");
                        ctx.Result.Exportados++;
                    }
                    catch
                    {
                        ctx.DadosParaPDF.Add((nomePastaExibicao, categoria, "ERROR", true));
                        ctx.LogContent.AppendLine($"[ERROR] Failed to move {nomePastaExibicao}");
                        ctx.Result.Erros++;
                    }
                }
                else
                {
                    string dirName = Path.GetFileName(dir);
                    ctx.DadosParaPDF.Add((dirName, "Unknown", "NO INF", true));
                    ctx.LogContent.AppendLine($"[IGNORED - NO INF] {dirName}");
                    ctx.Result.SemInf++;
                    ApagarDiretorioSeguro(dir);
                }
            }
        }

        private void Passo3_BackupGpuOffline(ContextoProcessamento ctx)
        {
            ctx.Token.ThrowIfCancellationRequested();
            ctx.Progresso.Report((GetMsgTraduzida(ctx.Idioma, ChaveMsg.CopiandoGpu), 0, true));

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
                    ctx.Token.ThrowIfCancellationRequested();
                    string nomePasta = Path.GetFileName(dir).ToLower();
                    if (prefixosWMI.Any(p => nomePasta.StartsWith(p)))
                    {
                        string destinoGPU = Path.Combine(ctx.DestinoRaiz, "GPU_DriverStore_" + Path.GetFileName(dir));
                        try
                        {
                            List<string> lockedFiles = new List<string>();
                            CopiarPastaCompleta(dir, destinoGPU, ctx.Token, lockedFiles);
                            ctx.AdicionarEspaco(ObterTamanhoPasta(destinoGPU));

                            if (lockedFiles.Count == 0)
                            {
                                ctx.DadosParaPDF.Add((Path.GetFileName(destinoGPU), "GPU_Offline_Pack", "OK", false));
                                ctx.LogContent.AppendLine($"[GPU OK] Copied {Path.GetFileName(dir)}");
                            }
                            else
                            {
                                string lockedStr = string.Join(", ", lockedFiles.Take(3)) + (lockedFiles.Count > 3 ? "..." : "");
                                ctx.DadosParaPDF.Add((Path.GetFileName(destinoGPU), "GPU_Offline_Pack", $"OK ({lockedFiles.Count} Locked)", false));
                                ctx.LogContent.AppendLine($"[GPU PARTIAL] Copied {Path.GetFileName(dir)} ({lockedFiles.Count} locked: {lockedStr})");
                            }

                            ctx.Result.GpuPacks++;
                        }
                        catch (OperationCanceledException)
                        {
                            ApagarDiretorioSeguro(destinoGPU);
                            throw;
                        }
                        catch (Exception ex)
                        {
                            ctx.DadosParaPDF.Add((Path.GetFileName(destinoGPU), "GPU_Offline_Pack", "ERROR", true));
                            ctx.LogContent.AppendLine($"[GPU ERROR] Failed {Path.GetFileName(dir)}: {ex.Message}");
                            ctx.Result.Erros++;
                        }
                    }
                }
            }
        }

        private void Passo4_BackupImpressoras(ContextoProcessamento ctx)
        {
            ctx.Token.ThrowIfCancellationRequested();
            ctx.Progresso.Report((GetMsgTraduzida(ctx.Idioma, ChaveMsg.CopiandoImp), 0, true));

            string printerFile = Path.Combine(ctx.DestinoRaiz, "Backup_Impressoras.printerExport");
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
                            while (!pPrint.WaitForExit(500)) { ctx.Token.ThrowIfCancellationRequested(); }
                        }
                        catch (OperationCanceledException)
                        {
                            try { if (!pPrint.HasExited) pPrint.Kill(); } catch { }
                            throw;
                        }
                    }

                    if (File.Exists(printerFile))
                    {
                        ctx.DadosParaPDF.Add(("Backup_Impressoras.printerExport", "Printer_Config", "OK", false));
                        ctx.LogContent.AppendLine($"[PRINTER OK] Configuracoes Exportadas");
                        ctx.Result.StatusImpressora = "OK (.printerExport)";
                        ctx.AdicionarEspaco(new FileInfo(printerFile).Length);
                    }
                    else
                    {
                        ctx.DadosParaPDF.Add(("Backup_Impressoras.printerExport", "Printer_Config", "ERROR", true));
                        ctx.LogContent.AppendLine($"[PRINTER ERROR] Falha na geracao");
                        ctx.Result.StatusImpressora = "Erro";
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    ctx.DadosParaPDF.Add(("Backup_Impressoras.printerExport", "Printer_Config", "EXEC ERROR", true));
                    ctx.LogContent.AppendLine($"[PRINTER ERROR] Falha de execucao do PrintBrm: {ex.Message}");
                    ctx.Result.StatusImpressora = "Erro Execução";
                }
            }
        }

        private void Passo5_GerarScriptsBat(ContextoProcessamento ctx, bool modoGPU, bool modoImp)
        {
            ctx.Token.ThrowIfCancellationRequested();

            foreach (var kvp in ctx.PastasPorCategoria)
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
                string pathBat = Path.Combine(ctx.DestinoRaiz, $"Instalar_Categoria_{cat}.bat");
                File.WriteAllText(pathBat, sbBat.ToString(), Encoding.Default);
                ctx.AdicionarEspaco(new FileInfo(pathBat).Length);
            }

            if (ctx.Result.GpuPacks > 0)
            {
                string batPathGPU = Path.Combine(ctx.DestinoRaiz, "0_Instalar_GPU_Offline.bat");
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
                ctx.AdicionarEspaco(new FileInfo(batPathGPU).Length);
            }

            if (ctx.Result.Exportados > 0 || ctx.Result.GpuPacks > 0)
            {
                StringBuilder sbMaster = new StringBuilder();
                sbMaster.AppendLine("@echo off\r\ntitle OmniDriver - Instalacao Completa");
                sbMaster.AppendLine("net session >nul 2>&1 || (echo ERRO: Execute Administrador! & pause & exit)");
                sbMaster.AppendLine("for /D %%G in (\"%~dp0*\") do (");
                sbMaster.AppendLine("    echo %%~nxG | findstr /I /B \"GPU_\" >nul || (");
                sbMaster.AppendLine("        if exist \"%%G\\*.inf\" ( pnputil /add-driver \"%%G\\*.inf\" /install /subdirs >nul )");
                sbMaster.AppendLine("    )\r\n)");
                if (ctx.Result.GpuPacks > 0)
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
                string pathMaster = Path.Combine(ctx.DestinoRaiz, "1_Instalar_TODOS_OS_DRIVERS.bat");
                File.WriteAllText(pathMaster, sbMaster.ToString(), Encoding.Default);
                ctx.AdicionarEspaco(new FileInfo(pathMaster).Length);
            }
        }

        private void Passo6_FinalizarRelatorios(ContextoProcessamento ctx, string caminhoPDF, string caminhoLog)
        {
            ctx.Token.ThrowIfCancellationRequested();
            ctx.Progresso.Report((GetMsgTraduzida(ctx.Idioma, ChaveMsg.GerandoPdf), 0, true));

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

                    document.Add(new Paragraph($"PC: {Environment.MachineName} | Date: {DateTime.Now}\n")
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).SetFontSize(10));

                    string hwTitle = ctx.Idioma == "PT" ? "Especificações do Sistema:" : ctx.Idioma == "EN" ? "System Specifications:" : "Especificaciones del Sistema:";
                    document.Add(new Paragraph(hwTitle).SetFont(fontBold).SetFontSize(11).SetFontColor(ColorConstants.DARK_GRAY));
                    document.Add(new Paragraph(ctx.Result.HardwareInfo).SetFont(fontMono).SetFontSize(9).SetMarginBottom(10));

                    Table table = new Table(UnitValue.CreatePercentArray(new float[] { 50, 30, 20 })).UseAllAvailableWidth();
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Folder Name / HW ID").SetFont(fontBold)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Category").SetFont(fontBold)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Status").SetFont(fontBold)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));

                    foreach (var d in ctx.DadosParaPDF)
                    {
                        table.AddCell(new Cell().Add(new Paragraph(d.Pasta).SetFontSize(9)));
                        table.AddCell(new Cell().Add(new Paragraph(d.Cat).SetFontSize(8)));

                        var color = d.Erro ? ColorConstants.RED : (d.Status.Contains("NO INF") ? ColorConstants.ORANGE : ColorConstants.GREEN);
                        table.AddCell(new Cell().Add(new Paragraph(d.Status).SetFontSize(8).SetFontColor(color).SetFont(fontBold)));
                    }
                    document.Add(table);
                }
                if (File.Exists(caminhoPDF)) ctx.AdicionarEspaco(new FileInfo(caminhoPDF).Length);
            }
            catch (Exception exPdf)
            {
                ctx.Result.StatusPdf = "Falha";
                ctx.LogContent.AppendLine($"[PDF ERROR] {exPdf.Message}");
            }

            ctx.Result.TamanhoFormatado = (ctx.EspacoUsadoBytes / (1024.0 * 1024.0)).ToString("0.00") + " MB";
            if (ctx.EspacoUsadoBytes > 1024 * 1024 * 1024) ctx.Result.TamanhoFormatado = (ctx.EspacoUsadoBytes / (1024.0 * 1024.0 * 1024.0)).ToString("0.00") + " GB";

            string finalResumo = $"[RESUMO FINAL]\r\nDrivers Exportados: {ctx.Result.Exportados}\r\nPacotes GPU: {ctx.Result.GpuPacks}\r\nSem INF: {ctx.Result.SemInf}\r\nErros: {ctx.Result.Erros}\r\nEspaco Usado: {ctx.Result.TamanhoFormatado}\r\nPDF: {ctx.Result.StatusPdf}\r\n\r\n";

            try
            {
                File.WriteAllText(caminhoLog, finalResumo + ctx.LogContent.ToString(), Encoding.Default);
                ctx.Result.StatusLog = "OK";
            }
            catch (Exception exLog)
            {
                ctx.Result.StatusLog = $"Falha: {exLog.Message}";
            }
        }

        // ==========================================
        // HELPERS GENERALES
        // ==========================================

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

        private void PrepararUI(bool trabalhando, string idioma, ModoOperacao modo)
        {
            _isProcessando = trabalhando;
            btnBrowse.Enabled = !trabalhando;
            cmbLanguage.Enabled = !trabalhando;

            btnExtrair.Enabled = !trabalhando;
            if (btnRestaurar != null) btnRestaurar.Enabled = !trabalhando;

            if (trabalhando)
            {
                string txt = idioma == "PT" ? "Cancelar" : idioma == "EN" ? "Cancel" : "Cancelar";
                if (modo == ModoOperacao.Extrair)
                {
                    btnExtrair.Text = txt;
                    btnExtrair.Enabled = true;
                }
                else if (modo == ModoOperacao.Restaurar && btnRestaurar != null)
                {
                    btnRestaurar.Text = txt;
                    btnRestaurar.Enabled = true;
                }
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
                case ChaveMsg.SemDrivers: return idioma == "PT" ? "O sistema não contém drivers de terceiros para exportar." : idioma == "EN" ? "System contains no third-party drivers to export." : "El sistema no contiene controladores de terceros para exportar.";
                default:
                    System.Diagnostics.Debug.Assert(false, $"ChaveMsg não tratada: {chave}");
                    return chave.ToString();
            }
        }

        // Formateo seguro para la UI de Restauración (evita FormatException)
        private string GetMsgInstalando(string idioma, string nomePasta)
        {
            string template = idioma == "PT" ? "A instalar: {0}..." : idioma == "EN" ? "Installing: {0}..." : "Instalando: {0}...";
            string safeName = nomePasta.Replace("{", "{{").Replace("}", "}}");
            return string.Format(template, safeName);
        }
    }
}