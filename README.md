# SkipBo Event Automation Tool

A specialized Unity Editor extension designed to automate the repetitive tasks of event creation and reopening for the SkipBo mobile game.

## 🚀 Key Features
- **Centralized Pipeline**: Executes 21 critical steps (Proto injection, Excel cloning, Localization updates, etc.) through a single GUI.
- **Intelligent Synchronization**:
  - Automatically remaps **Item IDs** between `item.xlsx` and `events.xlsx`.
  - Duplicates and renames **Icon Items** (Token/Material) based on template rows 69-70.
  - Updates **Sprite Names** using 8-digit Event IDs.
- **Safety First**: Automatically creates backups (`.bak`) before modifying any Excel file.
- **Coder Convert Integration**: Triggers the standard `coder_convert` tool as part of the final validation step.

## 🛠️ Installation
1. Copy the `EventAutomation` folder into your Unity project's `Assets/Editor/` directory.
2. Ensure you have **Python 3.x** installed with the following dependencies:
   ```bash
   pip install openpyxl
   ```
3. Open the tool in Unity via **Tools > Event Automator**.

## 📖 Usage
1. **Source Event**: Enter the name of the previous event to clone from (e.g., `Sim2507`).
2. **Target Event**: Enter the name of the new event (e.g., `Sim2605`).
3. **Workspace Path**: Provide the absolute path to your project root.
4. **New Event Settings**: For new events, fill in the Event ID (8 digits) and scheduling times.
5. **Execute**: Click "Execute" to run the full pipeline, or run individual steps manually if needed.

## 📁 Repository Structure
- `EventAutomationWindow.cs`: The Unity UI logic and process manager.
- `PythonBackend/event_pipeline.py`: The core automation logic.
- `EventAutomationWindow.uxml/uss`: UI Layout and Styling.
