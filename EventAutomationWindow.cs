using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;
using UnityEditor;
using UnityEngine;
using UnityEngine.UIElements;
using UnityEditor.UIElements;

public class EventAutomationWindow : EditorWindow
{
    private TextField sourceEventInput;
    private TextField targetEventInput;
    private TextField workspacePathInput;
    private TextField eventIdInput;
    private TextField startTimeInput;
    private TextField endTimeInput;
    private TextField nearEndTimeInput;
    private TextField closeTimeInput;
    private VisualElement newEventSettings;
    private RadioButton reopenRadio;
    private RadioButton newRadio;
    private Button executeBtn;
    private Button viewLogsBtn;
    private Button undoAllBtn;
    private ScrollView checklistContainer;
    private ScrollView consoleScrollView;

    private readonly List<string> pipelineSteps = new List<string>
    {
        "1. Configure storagedata.proto",
        "2. Execute ProtoToLua in Unity",
        "3. Add active save file in local service",
        "4. Register in LocalServiceMgr.lua",
        "5. Clone WorkSpace3 configuration to Design",
        "6. Clone Event Excel configuration tables",
        "7. Copy script to app_client directory",
        "8. Modify convert_layout.lua & convert.json",
        "9. Update bi.xlsx (BI tracking)",
        "10. Update events.xlsx (Event registration)",
        "11. Update event_shop.xlsx (Shop config)",
        "12. Update item.xlsx (Item definitions)",
        "13. Update icon.xlsx (Icon references)",
        "14. Update localization, quiz & answer_challenge",
        "15. Update asset_ref.xlsx (Asset references)",
        "16. Update store.xlsx (Store config)",
        "17. Update card.xlsx (Card config)",
        "18. Update pack.xlsx (Pack config)",
        "19. Update guide.xlsx (Guide config)",
        "20. Update sys.xlsx (DeepLink registration)",
        "21. Execute coder_convert script"
    };

    private List<VisualElement> uiStepItems = new List<VisualElement>();
    private bool scrollScheduled = false;

    [MenuItem("Tools/Event Automation Tool")]
    public static void ShowExample()
    {
        EventAutomationWindow wnd = GetWindow<EventAutomationWindow>();
        wnd.titleContent = new GUIContent("Event Automator");
        wnd.minSize = new Vector2(600, 700);
    }

    public void CreateGUI()
    {
        // Import UXML
        var visualTree = AssetDatabase.LoadAssetAtPath<VisualTreeAsset>("Assets/Editor/EventAutomation/EventAutomationWindow.uxml");
        if (visualTree == null)
        {
            UnityEngine.Debug.LogError("Could not load EventAutomationWindow.uxml");
            return;
        }
        
        VisualElement labelFromUXML = visualTree.Instantiate();
        rootVisualElement.Add(labelFromUXML);

        // Bind UI Elements
        sourceEventInput = rootVisualElement.Q<TextField>("source-event-input");
        targetEventInput = rootVisualElement.Q<TextField>("target-event-input");
        workspacePathInput = rootVisualElement.Q<TextField>("workspace-path-input");
        eventIdInput = rootVisualElement.Q<TextField>("event-id-input");
        startTimeInput = rootVisualElement.Q<TextField>("start-time-input");
        endTimeInput = rootVisualElement.Q<TextField>("end-time-input");
        nearEndTimeInput = rootVisualElement.Q<TextField>("near-end-time-input");
        closeTimeInput = rootVisualElement.Q<TextField>("close-time-input");
        newEventSettings = rootVisualElement.Q<VisualElement>("new-event-settings");
        reopenRadio = rootVisualElement.Q<RadioButton>("radio-reopen");
        newRadio = rootVisualElement.Q<RadioButton>("radio-new");

        executeBtn = rootVisualElement.Q<Button>("execute-btn");
        viewLogsBtn = rootVisualElement.Q<Button>("view-logs-btn");
        undoAllBtn = rootVisualElement.Q<Button>("undo-all-btn");
        checklistContainer = rootVisualElement.Q<ScrollView>("checklist-container");
        consoleScrollView = rootVisualElement.Q<ScrollView>("console-scrollview");

        reopenRadio.RegisterValueChangedCallback(evt => OnModeChanged());
        newRadio.RegisterValueChangedCallback(evt => OnModeChanged());
        
        // Initial state
        OnModeChanged();

        // Init Checklist
        BuildChecklistUI();

        // Bind Actions
        executeBtn.clicked += OnExecuteClicked;
        if (viewLogsBtn != null) 
            viewLogsBtn.clicked += ViewLogsClicked;
        if (undoAllBtn != null)
            undoAllBtn.clicked += OnUndoAllClicked;

        Log("Tool Initialized. Waiting for input...");
    }

    private void ViewLogsClicked()
    {
        Type consoleWindowType = typeof(EditorWindow).Assembly.GetType("UnityEditor.ConsoleWindow");
        if (consoleWindowType != null)
        {
            EditorWindow.GetWindow(consoleWindowType).Show();
        }
    }

    private void OnModeChanged()
    {
        var sourceRow = rootVisualElement.Q<VisualElement>("source-row");
        var targetRow = rootVisualElement.Q<VisualElement>("target-row");
        var sourceLabel = rootVisualElement.Q<Label>("source-label");

        if (reopenRadio.value)
        {
            // Reopen mode
            sourceLabel.text = "Event Name";
            sourceEventInput.ElementAt(0).tooltip = "The event you want to reopen";
            targetRow.style.display = DisplayStyle.None;
            if (newEventSettings != null)
                newEventSettings.style.display = DisplayStyle.None;
        }
        else
        {
            // New Event mode
            sourceLabel.text = "Source Event Name";
            sourceEventInput.ElementAt(0).tooltip = "The old event to copy from (e.g. merge2603)";
            targetRow.style.display = DisplayStyle.Flex;
            if (newEventSettings != null)
                newEventSettings.style.display = DisplayStyle.Flex;
        }
    }

    private void BuildChecklistUI()
    {
        checklistContainer.Clear();
        uiStepItems.Clear();

        for (int i = 0; i < pipelineSteps.Count; i++)
        {
            var itemContainer = new VisualElement();
            itemContainer.AddToClassList("checklist-item");

            var statusIcon = new VisualElement();
            statusIcon.AddToClassList("checklist-item-status-icon");
            statusIcon.AddToClassList("status-pending"); // Default class
            
            var textLabel = new Label(pipelineSteps[i]);
            textLabel.AddToClassList("checklist-item-text");

            var runBtn = new Button();
            runBtn.text = "Run";
            runBtn.AddToClassList("checklist-item-btn");

            var undoBtn = new Button();
            undoBtn.text = "Undo";
            undoBtn.AddToClassList("checklist-item-undo-btn");
            undoBtn.tooltip = "Revert this step";
            
            int stepIndex = i; // Closure capture
            runBtn.clicked += () => _ = ExecuteSingleStepAsync(stepIndex, false);
            undoBtn.clicked += () => _ = ExecuteSingleStepAsync(stepIndex, true);

            itemContainer.Add(statusIcon);
            itemContainer.Add(textLabel);
            itemContainer.Add(runBtn);
            itemContainer.Add(undoBtn);

            checklistContainer.Add(itemContainer);
            uiStepItems.Add(itemContainer);
        }
    }

    private void UpdateStepStatus(int index, string statusClass)
    {
        if (index < 0 || index >= uiStepItems.Count) return;

        var icon = uiStepItems[index].Q<VisualElement>(className: "checklist-item-status-icon");
        
        // Remove all previous status classes
        icon.RemoveFromClassList("status-pending");
        icon.RemoveFromClassList("status-running");
        icon.RemoveFromClassList("status-success");
        icon.RemoveFromClassList("status-error");

        // Add new class
        icon.AddToClassList(statusClass);
    }

    private void Log(string msg)
    {
        string timestamp = DateTime.Now.ToString("HH:mm:ss");
        if (consoleScrollView != null)
        {
            var logLine = new Label($"[{timestamp}] {msg}");
            logLine.enableRichText = true;
            
            if (msg.Contains("<color=red>") || msg.Contains("[CoderConvert Err]") || msg.Contains("Error") || msg.Contains("Failed"))
            {
                logLine.AddToClassList("console-log-error");
            }
            else
            {
                logLine.AddToClassList("console-log-line");
            }
            
            consoleScrollView.contentContainer.Add(logLine);
            
            // Limit line count to preserve high UI performance 
            if (consoleScrollView.contentContainer.childCount > 400)
            {
                consoleScrollView.contentContainer.RemoveAt(0);
            }
            
            // Auto scroll to latest output safely with debouncing
            if (!scrollScheduled)
            {
                scrollScheduled = true;
                EditorApplication.delayCall += () => {
                    scrollScheduled = false;
                    if (consoleScrollView != null && consoleScrollView.contentContainer.childCount > 0)
                    {
                        var lastChild = consoleScrollView.contentContainer[consoleScrollView.contentContainer.childCount - 1];
                        consoleScrollView.ScrollTo(lastChild);
                    }
                };
            }
        }
        UnityEngine.Debug.Log($"[EventAutomation] {msg}");
    }

    private void OnExecuteClicked()
    {
        if (reopenRadio.value)
        {
            if (string.IsNullOrEmpty(sourceEventInput.value))
            {
                Log("<color=red>Error: Event Name cannot be empty.</color>");
                EditorUtility.DisplayDialog("Validation Error", "Please fill in all required fields.", "OK");
                return;
            }
            targetEventInput.value = sourceEventInput.value; // For backend parity, treat target as same as source
        }
        else
        {
            if (string.IsNullOrEmpty(sourceEventInput.value) || string.IsNullOrEmpty(targetEventInput.value))
            {
                Log("<color=red>Error: Source and Target Event Names cannot be empty.</color>");
                EditorUtility.DisplayDialog("Validation Error", "Please fill in all required fields.", "OK");
                return;
            }
        }

        string mode = reopenRadio.value ? "Reopen" : "New Event";
        Log($"Starting full pipeline. Mode: {mode} | Source: {sourceEventInput.value} | Target: {targetEventInput.value}");

        // For now, we simulate the execution process using an Editor Coroutine or Task
        // We'll use a simple background execution for the prototype.
        StartMockExecution();
    }

    private bool isExecuting = false;
    private bool isRevertMode = false;

    private async void StartMockExecution()
    {
        executeBtn.SetEnabled(false);
        if (undoAllBtn != null) undoAllBtn.SetEnabled(false);
        Log("Execution started...");
        
        isExecuting = true;
        isRevertMode = false;
        
        for (int i = 0; i < pipelineSteps.Count; i++)
        {
            if (!isExecuting) break;
            bool success = await ExecuteSingleStepAsync(i, false);
            if (!success) {
                Log("Pipeline aborted due to step failure.");
                break;
            }
        }
        
        isExecuting = false;
        executeBtn.SetEnabled(true);
        if (undoAllBtn != null) undoAllBtn.SetEnabled(true);
        Log("Execution process finished.");
    }
    
    private async void OnUndoAllClicked()
    {
        executeBtn.SetEnabled(false);
        if (undoAllBtn != null) undoAllBtn.SetEnabled(false);
        Log("<color=#FFEA00>Revert Pipeline started...</color>");
        
        isExecuting = true;
        isRevertMode = true;
        
        // Undo happens in reverse order
        for (int i = pipelineSteps.Count - 1; i >= 0; i--)
        {
            if (!isExecuting) break;
            await ExecuteSingleStepAsync(i, true);
        }

        isExecuting = false;
        executeBtn.SetEnabled(true);
        if (undoAllBtn != null) undoAllBtn.SetEnabled(true);
        Log("Revert process finished.");
    }

    private async Task<bool> ExecuteSingleStepAsync(int index, bool isRevert)
    {
        string opStr = isRevert ? "Reverting" : "Executing";
        Log($"{opStr} Single Step {index + 1}: {pipelineSteps[index]}...");
        UpdateStepStatus(index, "status-running");
        
        bool success = await ProcessStepLogicAsync(index, isRevert);
        
        if (success)
        {
            UpdateStepStatus(index, isRevert ? "status-pending" : "status-success");
            Log(isRevert ? $"<color=#00E676>Step {index + 1} Reverted.</color>" : $"<color=#00E676>Step {index + 1} Succeeded.</color>");
            return true;
        }
        else
        {
            UpdateStepStatus(index, "status-error");
            Log($"<color=#FF1744>Step {index + 1} Failed!</color>");
            return false;
        }
    }

    private async Task<bool> ProcessStepLogicAsync(int stepIndex, bool isRevert)
    {
        string srcName = sourceEventInput.value;
        string tgtName = targetEventInput.value;
        
        if (reopenRadio.value)
        {
            tgtName = srcName;
        }

        if (string.IsNullOrEmpty(srcName))
        {
            Log("<color=red>Error: Source Event Name cannot be empty.</color>");
            return false;
        }

        if (!reopenRadio.value && string.IsNullOrEmpty(tgtName))
        {
            Log("<color=red>Error: Target Event Name cannot be empty.</color>");
            return false;
        }

        if (stepIndex == 1)
        {
            if (isRevert)
            {
                Log("Reverting ProtoToLua is not strictly trackable. Skipping.");
                return true;
            }
            try 
            {
                Log("Executing native AssetBundleTools/ProtoToLua menu item...");
                ProtoTool.ProtoToLua();
                Log("ProtoToLua Execution Complete.");
                return true;
            }
            catch(Exception e)
            {
                Log($"<color=red>ProtoToLua Failed: {e.Message}</color>");
                return false;
            }
        }

        string repoRoot = Path.GetFullPath(Path.Combine(Application.dataPath, "../../../../")).TrimEnd('\\', '/'); 

        if (stepIndex == 20)
        {
            if (isRevert)
            {
                Log("Reverting coder_convert is not supported directly. Skipping.");
                return true;
            }
            try
            {
                Log("Executing coder_convert_not_pause.bat...");
                string batPath = Path.Combine(repoRoot, @"design\DesignData\WorkSpcae_Design\coder_convert_not_pause.bat");
                if (!File.Exists(batPath))
                {
                    Log($"<color=red>Error: coder_convert_not_pause.bat not found at {batPath}</color>");
                    return false;
                }

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = batPath;
                startInfo.WorkingDirectory = Path.GetDirectoryName(batPath);
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                using (Process process = Process.Start(startInfo))
                {
                    process.OutputDataReceived += (sender, args) => {
                        if (!string.IsNullOrEmpty(args.Data))
                        {
                            string msg = args.Data;
                            EditorApplication.delayCall += () => Log($"[CoderConvert] {msg}");
                        }
                    };
                    process.ErrorDataReceived += (sender, args) => {
                        if (!string.IsNullOrEmpty(args.Data))
                        {
                            string msg = args.Data;
                            EditorApplication.delayCall += () => Log($"<color=red>[CoderConvert Err]</color> {msg}");
                        }
                    };
                    
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    await Task.Run(() => process.WaitForExit());

                    if (process.ExitCode != 0)
                    {
                        Log($"<color=red>CoderConvert Failed with exit code {process.ExitCode}</color>");
                        return false;
                    }
                }
                Log("CoderConvert Execution Complete.");
                return true;
            }
            catch (Exception e)
            {
                Log($"<color=red>CoderConvert Exception: {e.Message}</color>");
                return false;
            }
        }


        try 
        {
            string pyScriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/event_pipeline.py");
            
            if (!File.Exists(pyScriptPath))
            {
                Log($"<color=red>Error: Python script not found at {pyScriptPath}</color>");
                return false;
            }

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "cmd.exe"; 
            string revertArg = isRevert ? "1" : "0";
            string logFile = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/pipeline_log.txt");
            
            // Write extra args to a JSON file to avoid cmd escaping issues
            string extraArgsFile = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/extra_args.json");
            string eventId = eventIdInput != null ? eventIdInput.value : "";
            string startTime = startTimeInput != null ? startTimeInput.value : "";
            string endTime = endTimeInput != null ? endTimeInput.value : "";
            string nearEndTime = nearEndTimeInput != null ? nearEndTimeInput.value : "";
            string closeTime = closeTimeInput != null ? closeTimeInput.value : "";
            string isReopenStr = reopenRadio.value ? "1" : "0";
            string extraJson = $"{{\"event_id\":\"{eventId}\",\"start_time\":\"{startTime}\",\"end_time\":\"{endTime}\",\"near_end_time\":\"{nearEndTime}\",\"close_time\":\"{closeTime}\",\"is_reopen\":\"{isReopenStr}\"}}";
            Log($"Writing extra_args.json: {extraJson}");
            File.WriteAllText(extraArgsFile, extraJson);
            
            // Use cmd /c to output directly to a file, bypassing Unity Mono stream deadlocks
            startInfo.Arguments = $"/c python \"{pyScriptPath}\" {stepIndex} \"{srcName}\" \"{tgtName}\" \"{repoRoot}\" {revertArg} > \"{logFile}\" 2>&1";
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;

            using (Process process = Process.Start(startInfo))
            {
                await Task.Run(() => process.WaitForExit());

                string output = File.Exists(logFile) ? File.ReadAllText(logFile, System.Text.Encoding.UTF8) : "No Log Generated";

                if (!string.IsNullOrEmpty(output)) Log(output);
                
                if (process.ExitCode != 0)
                {
                    Log($"<color=red>Python Error: Script failed with code {process.ExitCode}</color>");
                    return false;
                }
            }
            
            return true;
        }
        catch (Exception ex)
        {
            Log($"<color=red>Exception during Step {stepIndex + 1}: {ex.Message}</color>");
            return false;
        }
    }
}
