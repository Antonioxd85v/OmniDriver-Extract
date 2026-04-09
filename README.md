⚙️ OmniDriver Extract
is a professional, standalone Windows utility designed for IT technicians, system administrators, and power users. It provides a highly reliable way to safely export all system drivers and automatically generates a detailed PDF audit report of the hardware configuration.

Whether you are preparing to format a PC, creating offline backups, or conducting an enterprise hardware audit, OmniDriver Extract ensures you never lose a critical driver again.

✨ Key Features
🚀 Comprehensive Extraction: Utilizes native Windows deployment engines (DISM) to safely back up all functional and third-party drivers without missing complex components like Chipsets or Bluetooth modules.

📂 Smart Organization: Say goodbye to confusing generic folder names (like oem12.inf). The application intelligently scans the INF files and names the output folders based on the Manufacturer, Category, or Hardware ID (e.g., AMD_Chipset, Realtek_Network, Intel_System).

📄 Automated PDF Audit Reports: Automatically generates a clean, professional PDF document mapping every extracted driver to its specific folder. The report opens instantly upon completion.

🛡️ Safe Privilege Escalation: Designed with modern Windows security in mind. The app launches as a standard user to avoid SmartScreen blocks, only requesting UAC Administrator privileges at the exact moment of extraction.

🌍 Multilingual Interface: Fully localized user interface supporting English, Spanish, and Portuguese, making it accessible to a global audience.

⚡ Zero Dependencies (Standalone): Compiled as a self-contained, single-file executable. It requires no .NET runtime installations on the client's machine and runs flawlessly on both x86 and x64 architectures.

🛠️ How to Use
Run the Application: Simply double-click the installer or the portable executable.

Select Destination: Choose a folder on your local drive or an external USB drive where you want to save the backup.

Export: Click the "Export Drivers (PDF)" button. The application will ask for administrative permissions to securely access the Windows driver store.

Done! Wait a few minutes. Once finished, the tool will automatically open your PDF report, and your drivers will be perfectly organized and ready for a fresh Windows installation.

💻 Technical Stack
Language: C# / .NET

UI Framework: Windows Forms (WinForms)

PDF Engine: iText7 (with BouncyCastle Adapter)

Core Extraction: Windows DISM & PnPUtil APIs

Created by António — Providing smart solutions for IT professionals.
