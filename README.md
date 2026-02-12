# 🎤 AI PowerPoint Generator

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![LM Studio](https://img.shields.io/badge/Integration-LM%20Studio-orange.svg)](https://lmstudio.ai/)

<div align="center">
  <img src="docs/images/banner.png" alt="AI PowerPoint Generator Banner" width="600px">
  <p><em>Transform simple titles into complete PowerPoint presentations using local LLMs</em></p>
</div>

---

## 📋 Overview

**AI PowerPoint Generator** is a PowerShell script with a Windows Forms GUI that connects to a local LM Studio instance to automatically generate complete PowerPoint presentations. Simply enter a title and author name, and the AI handles everything:

- 📝 **Generates** a structured presentation outline
- 🎯 **Creates** 5-7 bullet points for each section
- 📊 **Builds** a professional PowerPoint slide deck
- 💾 **Saves** the finished presentation to your desktop

All AI processing runs **locally** on your machine through LM Studio - no API keys, no internet required, no privacy concerns.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🖥️ **Native Windows GUI** | Clean, responsive Windows Forms interface |
| 🤖 **Local AI Processing** | Works with LM Studio, completely offline |
| 📊 **PowerPoint Automation** | Full COM object integration with PowerPoint |
| 🔄 **Automatic Fallbacks** | Graceful degradation if AI service is unavailable |
| ⚡ **Real-time Status** | Live progress updates in status bar |
| 🎨 **Clean Formatting** | Properly formatted bullet points and titles |
| 📁 **Auto-save** | Timestamped files saved to desktop |
| 🛡️ **Error Handling** | Comprehensive error catching and user feedback |

---

## 🚀 Quick Start

### Prerequisites

1. **LM Studio** installed and running with server mode enabled
2. **Microsoft PowerPoint** (2013 or later recommended)
3. **PowerShell 5.1+** (comes pre-installed with Windows 10/11)
4. **Model loaded** in LM Studio (tested with `liquid/lfm2-1.2b`)

### Installation

```powershell
# Clone the repository
git clone https://github.com/Sba-Stuff/ai-powerpoint-generator.git

# Navigate to directory
cd ai-powerpoint-generator

# Run the script (you may need to bypass execution policy)
.\AI-PowerPoint-Generator.ps1
```

### Onliner Quick Run
Copy, paste in notepad. Save this as bat file along with the powershell script. Double click bat file to run.
```powershell
powershell -ExecutionPolicy Bypass -File "AI-PowerPoint-Generator.ps1"
```

### Things to Make Sure
1. Open LM Studio
2. Load your model (e.g., liquid/lfm2-1.2b)
3. Modify variables (ip, model, maxtokens, temperature) and prompt based on your choice.
4. Click "Start Server" button
5. Verify server runs (either at localhost:1234, or your set server IP. In this examplple, it is at http://192.168.10.4:1234)
