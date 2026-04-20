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
using System.Linq;

[Serializable]
public class MinigameInfo
{
    public string name;
    public int discount_id;
    public bool requires_double_week;
}

[Serializable]
public class MinigamesData
{
    public int max_id;
    public MinigameInfo[] minigames;
}

[Serializable]
public class BattlePassInfo
{
    public int last_cycle_id;
    public string last_bp_plan_alias;
    public string last_recharge_alias;
    public int last_item_id;
    public int last_scheme_id;
    public int last_800_cycle;
    public int last_900_cycle;
}

[Serializable]
public class HolidayBPInfo
{
    public int last_cycle_id;
    public int last_recommend_id;
    public int last_item_id;
    public int last_scheme_id;
    public string last_holiday_id_str;
    public string[] last_suffixes;
    public int last_title_set;
}

public class EventAutomationWindow : EditorWindow
{
    private TextField sourceEventInput;
    private TextField targetEventInput;
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

    // View Management
    private ScrollView mainEventsView;
    private ScrollView minigamesView;
    private VisualElement placeholderView;
    private Label placeholderTitle;
    private Label placeholderDesc;
    // Holiday BP UI
    private ScrollView holidayBPView;
    private TextField hbpHolidayId;
    private TextField hbpStartTime;
    private TextField hbpEndTime;
    private Label hbpNextCycle;
    private Label hbpNextRecommend;
    private Label hbpNextItem;
    private Label hbpNextScheme;
    private Label hbpSuffixPreview;
    private Button hbpExecuteBtn;
    private HolidayBPInfo _cachedHBPData;

    // Settings UI
    private ScrollView settingsView;
    private DropdownField workspaceDropdown;
    private Label resolvedPathLabel;
    private const string WorkspacePrefKey = "EventAutomation_SelectedWorkspace";
    private readonly List<string> WorkspaceOptions = new List<string> 
    { 
        "WorkSpcae_Design", 
        "WorkSpcae_NE", 
        "WorkSpcae_FC", 
        "WorkSpcae3" 
    };

    private List<Button> menuButtons = new List<Button>();

    // --- Minigame Specific Handles ---
    private VisualElement minigameRowsContainer;
    private Button addMinigameRowBtn;
    private Button minigameExecuteBtn;
    private Button minigameUndoBtn;
    private MinigamesData _cachedMinigamesData = null;

    // --- BattlePass Specific Handles ---
    private ScrollView battlepassView;
    private TextField bpStartTime;
    private TextField bpEndTime;
    private Label bpPointMode;
    private Label bpPrevCycle;
    private Label bpPrevPlan;
    private Label bpPrevYYMM;
    private Label bpPrevTemplate;
    private Label bpPrevItem;
    private Button bpExecuteBtn;
    private BattlePassInfo _cachedBPData = null;

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

        // View binds
        mainEventsView = rootVisualElement.Q<ScrollView>("main-events-view");
        minigamesView = rootVisualElement.Q<ScrollView>("minigames-view");
        placeholderView = rootVisualElement.Q<VisualElement>("placeholder-view");
        placeholderTitle = placeholderView?.Q<Label>("placeholder-title");
        placeholderDesc = placeholderView?.Q<Label>("placeholder-desc");

        // Settings Binds
        settingsView = rootVisualElement.Q<ScrollView>("settings-view");
        workspaceDropdown = rootVisualElement.Q<DropdownField>("workspace-dropdown");
        resolvedPathLabel = rootVisualElement.Q<Label>("resolved-path-label");

        // Minigames Bindings
        minigameRowsContainer = rootVisualElement.Q<VisualElement>("minigame-rows-container");
        addMinigameRowBtn = rootVisualElement.Q<Button>("add-minigame-row-btn");
        minigameExecuteBtn = rootVisualElement.Q<Button>("minigame-execute-btn");
        minigameUndoBtn = rootVisualElement.Q<Button>("minigame-undo-btn");

        if (minigameExecuteBtn != null) minigameExecuteBtn.clicked += OnMinigameExecuteClicked;

        // BattlePass Bindings
        battlepassView = rootVisualElement.Q<ScrollView>("battlepass-view");
        bpStartTime = rootVisualElement.Q<TextField>("bp-start-time");
        bpEndTime = rootVisualElement.Q<TextField>("bp-end-time");
        bpPointMode = rootVisualElement.Q<Label>("bp-point-mode");
        bpPrevCycle = rootVisualElement.Q<Label>("bp-prev-cycle");
        bpPrevPlan = rootVisualElement.Q<Label>("bp-prev-plan");
        bpPrevYYMM = rootVisualElement.Q<Label>("bp-prev-yymm");
        bpPrevTemplate = rootVisualElement.Q<Label>("bp-prev-template");
        bpPrevItem = rootVisualElement.Q<Label>("bp-prev-item");
        bpExecuteBtn = rootVisualElement.Q<Button>("bp-execute-btn");

        if (bpStartTime != null) bpStartTime.RegisterValueChangedCallback(evt => UpdateBPPreview());
        if (bpEndTime != null) bpEndTime.RegisterValueChangedCallback(evt => UpdateBPPreview());
        if (bpExecuteBtn != null) bpExecuteBtn.clicked += OnBattlePassCommitClicked;

        // Holiday BP Bindings
        holidayBPView = rootVisualElement.Q<ScrollView>("holiday-bp-view");
        hbpHolidayId = rootVisualElement.Q<TextField>("hbp-holiday-id");
        hbpStartTime = rootVisualElement.Q<TextField>("hbp-start-time");
        hbpEndTime = rootVisualElement.Q<TextField>("hbp-end-time");
        hbpNextCycle = rootVisualElement.Q<Label>("hbp-next-cycle");
        hbpNextRecommend = rootVisualElement.Q<Label>("hbp-next-recommend");
        hbpNextItem = rootVisualElement.Q<Label>("hbp-next-item");
        hbpNextScheme = rootVisualElement.Q<Label>("hbp-next-scheme");
        hbpSuffixPreview = rootVisualElement.Q<Label>("hbp-suffix-preview");
        hbpExecuteBtn = rootVisualElement.Q<Button>("hbp-execute-btn");

        if (hbpHolidayId != null) hbpHolidayId.RegisterValueChangedCallback(evt => UpdateHolidayBPPreview());
        if (hbpExecuteBtn != null) hbpExecuteBtn.clicked += OnHolidayBPCommitClicked;

        // Setup Sidebar Buttons
        SetupMenuButton("menu-main-events", "Main Events", "main");
        SetupMenuButton("menu-minigames", "Minigame Configs", "minigame");
        SetupMenuButton("menu-pet-bp", "Pet Battlepass", "battlepass");
        SetupMenuButton("menu-holiday-bp", "Holiday Battlepass", "holidaypass");
        SetupMenuButton("menu-deco-shop", "Decoration Shop", "placeholder");
        SetupMenuButton("menu-streak", "Streak Challenge", "placeholder");
        SetupMenuButton("menu-endless", "Endless Discount", "placeholder");
        SetupMenuButton("menu-settings", "Settings", "settings"); reopenRadio.RegisterValueChangedCallback(evt => OnModeChanged());
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

    private void SetupMenuButton(string btnName, string label, string viewType)
    {
        var btn = rootVisualElement.Q<Button>(btnName);
        if (btn == null) return;
        
        menuButtons.Add(btn);
        btn.clicked += () => OnMenuClicked(btn, label, viewType);
    }

    private void OnMenuClicked(Button clickedBtn, string label, string viewType)
    {
        // Update active classes
        foreach (var btn in menuButtons)
        {
            btn.RemoveFromClassList("active-btn");
        }
        clickedBtn.AddToClassList("active-btn");

        // Hide all views first
        if (mainEventsView != null) mainEventsView.style.display = DisplayStyle.None;
        if (minigamesView != null) minigamesView.style.display = DisplayStyle.None;
        if (battlepassView != null) battlepassView.style.display = DisplayStyle.None;
        if (holidayBPView != null) holidayBPView.style.display = DisplayStyle.None;
        if (settingsView != null) settingsView.style.display = DisplayStyle.None;
        if (placeholderView != null) placeholderView.style.display = DisplayStyle.None;

        if (viewType == "main")
        {
            if (mainEventsView != null) mainEventsView.style.display = DisplayStyle.Flex;
        }
        else if (viewType == "minigame")
        {
            if (minigamesView != null) minigamesView.style.display = DisplayStyle.Flex;
            InitializeMinigamesViewAsync();
        }
        else if (viewType == "battlepass")
        {
            if (battlepassView != null) battlepassView.style.display = DisplayStyle.Flex;
            InitializeBattlePassViewAsync();
        }
        else if (viewType == "holidaypass")
        {
            if (holidayBPView != null) holidayBPView.style.display = DisplayStyle.Flex;
            InitializeHolidayBPViewAsync();
        }
        else if (viewType == "settings")
        {
            if (settingsView != null) settingsView.style.display = DisplayStyle.Flex;
            InitializeSettingsView();
        }
        else
        {
            if (placeholderView != null) 
            {
                placeholderView.style.display = DisplayStyle.Flex;
                int firstSpace = label.IndexOf(' ');
                string cleanLabel = (firstSpace >= 0 && firstSpace < label.Length - 1) ? label.Substring(firstSpace + 1).Trim() : label;
                if (placeholderTitle != null) placeholderTitle.text = cleanLabel;
            }
        }
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
                string batPath = Path.Combine(repoRoot, "design", "DesignData", GetActiveWorkspaceName(), "coder_convert_not_pause.bat");
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
            startInfo.Arguments = $"/c python \"{pyScriptPath}\" {stepIndex} \"{srcName}\" \"{tgtName}\" \"{repoRoot}\" {revertArg} \"{GetActiveWorkspaceName()}\" > \"{logFile}\" 2>&1";
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

    // --- Minigame Configurations Logic ---

    [Serializable]
    public class MinigameRowPayload
    {
        public int id;
        public string minigame;
        public string start_time;
        public string end_time;
        public string double_week_id;
        public int discount_id;
    }

    [Serializable]
    public class MinigamePayloadList
    {
        public List<MinigameRowPayload> rows = new List<MinigameRowPayload>();
    }

    private async void InitializeMinigamesViewAsync()
    {
        if (_cachedMinigamesData != null) return;

        Log("Loading Minigame configs from Excel...");
        string pythonPath = "python";
        string scriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/get_minigames_info.py");
        string workspace_path = Application.dataPath.Replace("/develop/client/Skipbo/Assets", "").Replace("\\develop\\client\\Skipbo\\Assets", "");

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            Arguments = $"\"{scriptPath}\" {GetPythonDesignPathArg()}",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using (Process process = Process.Start(startInfo))
        {
            string output = await process.StandardOutput.ReadToEndAsync();
            string error = await process.StandardError.ReadToEndAsync();
            process.WaitForExit();

            if (process.ExitCode != 0)
            {
                Log($"Error loading minigames: {error}");
                return;
            }

            Log($"RAW JSON: {output}");

            try
            {
                _cachedMinigamesData = JsonUtility.FromJson<MinigamesData>(output);
                Log($"Parsed ID: {_cachedMinigamesData.max_id}, Parsed Minigames: {(_cachedMinigamesData.minigames != null ? _cachedMinigamesData.minigames.Length : 0)}");
                Log($"Loaded {(_cachedMinigamesData.minigames != null ? _cachedMinigamesData.minigames.Length : 0)} minigames out of history.");
                AddMinigameRow();
            }
            catch (Exception ex)
            {
                Log($"Error parsing minigame json: {ex.Message}\nRaw JSON: {output}");
            }
        }
    }

    private void AddMinigameRow()
    {
        if (_cachedMinigamesData == null) 
        {
            Log("Minigame configurations are still loading or failed to load, please wait...");
            return;
        }

        // Compute next ID dynamically from existing row count — avoids stale counter drift
        int existingRows = minigameRowsContainer.childCount;
        int nextId = _cachedMinigamesData.max_id + existingRows + 1;

        VisualElement row = new VisualElement();
        row.AddToClassList("form-row");
        row.style.backgroundColor = new StyleColor(new Color(0, 0, 0, 0.3f));
        row.style.paddingTop = 15;
        row.style.paddingBottom = 15;
        row.style.paddingLeft = 15;
        row.style.paddingRight = 15;
        row.style.borderTopLeftRadius = 12;
        row.style.borderTopRightRadius = 12;
        row.style.borderBottomLeftRadius = 12;
        row.style.borderBottomRightRadius = 12;
        row.style.marginBottom = 15;
        row.style.flexDirection = FlexDirection.Row;
        row.style.flexWrap = Wrap.Wrap;
        row.style.alignItems = Align.Center;

        // ID label + editable input side-by-side
        Label idLabel = new Label("ID");
        idLabel.style.color = new StyleColor(new Color(1f, 0.85f, 0.2f));
        idLabel.style.unityFontStyleAndWeight = FontStyle.Bold;
        idLabel.style.alignSelf = Align.Center;
        idLabel.style.marginRight = 4;

        TextField idField = new TextField();
        idField.value = nextId.ToString();
        idField.style.width = 80;
        // Suppress the built-in label so the full width is the input box
        idField.labelElement.style.display = DisplayStyle.None;
        idField.labelElement.style.minWidth = 0;
        idField.labelElement.style.width = 0;
        
        DropdownField drop = new DropdownField("Minigame");
        drop.choices = _cachedMinigamesData.minigames.Select(m => m.name).ToList();
        drop.value = drop.choices.FirstOrDefault();
        drop.style.width = 200;
        
        TextField startField = new TextField("Start Time");
        startField.value = "2026-06-01 00:00:00";
        startField.style.width = 220;

        TextField endField = new TextField("End Time");
        endField.value = "2026-06-03 23:59:59";
        endField.style.width = 220;

        TextField dwField = new TextField("Double Wk ID");
        dwField.style.width = 150;
        dwField.style.display = DisplayStyle.None;

        drop.RegisterValueChangedCallback(evt =>
        {
            var mg = _cachedMinigamesData.minigames.FirstOrDefault(m => m.name == drop.value);
            dwField.style.display = (mg != null && mg.requires_double_week) ? DisplayStyle.Flex : DisplayStyle.None;
        });

        var initMg = _cachedMinigamesData.minigames.FirstOrDefault(m => m.name == drop.value);
        if (initMg != null && initMg.requires_double_week) dwField.style.display = DisplayStyle.Flex;

        Button delBtn = new Button(() => 
        { 
            minigameRowsContainer.Remove(row);
        }) { text = "✕" };
        delBtn.AddToClassList("danger-btn");
        delBtn.style.width = 40;
        delBtn.style.height = 36;
        delBtn.style.alignSelf = Align.Center;

        row.Add(idLabel);
        row.Add(idField);
        row.Add(drop);
        row.Add(startField);
        row.Add(endField);
        row.Add(dwField);
        row.Add(delBtn);

        minigameRowsContainer.Add(row);
    }

    private async void OnMinigameExecuteClicked()
    {
        // Build JSON manually — JsonUtility.ToJson corrupts non-ASCII (Chinese) characters to ???
        var sb = new System.Text.StringBuilder();
        sb.Append("{\"rows\":[");
        bool first = true;

        foreach (var child in minigameRowsContainer.Children())
        {
            // Row layout: [0]=ID Label, [1]=idField(TextField), [2]=drop, [3]=start, [4]=end, [5]=dw, [6]=delBtn
            TextField idField = child.ElementAt(1) as TextField;
            DropdownField drop = child.ElementAt(2) as DropdownField;
            TextField start = child.ElementAt(3) as TextField;
            TextField end = child.ElementAt(4) as TextField;
            TextField dw = child.ElementAt(5) as TextField;

            if (idField == null || drop == null) continue;

            int id;
            if (!int.TryParse(idField.value, out id)) 
            {
                Log($"Skipping row — invalid ID: '{idField.value}'");
                continue;
            }

            var mg = _cachedMinigamesData.minigames.FirstOrDefault(m => m.name == drop.value);
            int discountId = mg != null ? mg.discount_id : -1;
            string dwVal = (dw != null && dw.style.display == DisplayStyle.Flex) ? dw.value : "";

            // Escape fields for JSON safety
            string safeName = drop.value.Replace("\\", "\\\\").Replace("\"", "\\\"");
            string safeStart = start.value.Replace("\"", "\\\"");
            string safeEnd = end.value.Replace("\"", "\\\"");
            string safeDw = dwVal.Replace("\"", "\\\"");

            if (!first) sb.Append(",");
            sb.Append($"{{\"id\":{id},\"minigame\":\"{safeName}\",\"start_time\":\"{safeStart}\",\"end_time\":\"{safeEnd}\",\"double_week_id\":\"{safeDw}\",\"discount_id\":{discountId}}}");
            first = false;
        }
        sb.Append("]}");

        string json = sb.ToString();
        Log($"PAYLOAD JSON: {json}");
        
        Log("Sending minigame entries to python...");
        string pythonPath = "python";
        string scriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/update_minigames.py");
        string workspace_path = Application.dataPath.Replace("/develop/client/Skipbo/Assets", "").Replace("\\develop\\client\\Skipbo\\Assets", "");

        // Write payload as UTF8 so Chinese characters are NOT corrupted
        string tempJsonPath = Path.Combine(Path.GetTempPath(), "minigames_payload.json");
        File.WriteAllText(tempJsonPath, json, System.Text.Encoding.UTF8);
        Log($"Payload written to: {tempJsonPath}");

        // Disable commit button during processing
        if (minigameExecuteBtn != null) minigameExecuteBtn.SetEnabled(false);
        if (minigameExecuteBtn != null) minigameExecuteBtn.text = "Committing...";

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            Arguments = $"\"{scriptPath}\" {GetPythonDesignPathArg()} \"{tempJsonPath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using (Process process = Process.Start(startInfo))
        {
            string output = await process.StandardOutput.ReadToEndAsync();
            string error = await process.StandardError.ReadToEndAsync();
            process.WaitForExit();

            // Re-enable button
            if (minigameExecuteBtn != null)
            {
                minigameExecuteBtn.SetEnabled(true);
                minigameExecuteBtn.text = "Commit to mini_mgr.xlsx";
            }

            if (process.ExitCode != 0)
            {
                string errMsg = !string.IsNullOrEmpty(error) ? error : output;
                Log($"<color=red>Commit failed: {errMsg}</color>");
                EditorUtility.DisplayDialog(
                    "Commit Failed ❌",
                    string.IsNullOrEmpty(errMsg) ? "Unknown error. Check that mini_mgr.xlsx is closed in Excel." : errMsg,
                    "OK"
                );
            }
            else
            {
                int rowCount = minigameRowsContainer.childCount;
                Log($"<color=green>Success: {output}</color>");
                EditorUtility.DisplayDialog(
                    "Commit Successful ✅",
                    $"{rowCount} Minigame run(s) appended to mini_mgr.xlsx successfully!",
                    "Great!"
                );
                // Clear rows after successful commit
                minigameRowsContainer.Clear();
                // Reset cached data so next open re-reads updated max_id
                _cachedMinigamesData = null;
            }
        }
    }

    // --- BattlePass Configurations Logic ---

    private async void InitializeBattlePassViewAsync()
    {
        if (_cachedBPData != null) return;

        Log("Loading BattlePass metadata from Excel...");
        
        string pythonPath = "python";
        string scriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/get_bp_info.py");
        string workspaceRoot = Application.dataPath.Replace("/develop/client/Skipbo/Assets", "").Replace("\\develop\\client\\Skipbo\\Assets", "");

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            Arguments = $"\"{scriptPath}\" {GetPythonDesignPathArg()}",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using (Process process = Process.Start(startInfo))
        {
            string output = await process.StandardOutput.ReadToEndAsync();
            string error = await process.StandardError.ReadToEndAsync();
            process.WaitForExit();

            if (process.ExitCode == 0)
            {
                try
                {
                    _cachedBPData = JsonUtility.FromJson<BattlePassInfo>(output);
                    Log($"[EventAutomation] BP Metadata Loaded. Last Cycle: {_cachedBPData.last_cycle_id}");
                    UpdateBPPreview();
                }
                catch (Exception ex)
                {
                    Log($"<color=red>Error parsing BP JSON: {ex.Message}</color>");
                    UnityEngine.Debug.Log($"RAW BP JSON: {output}");
                }
            }
            else
            {
                Log($"<color=red>Error loading BP info: {error}</color>");
            }
        }
    }

    private void UpdateBPPreview()
    {
        if (_cachedBPData == null) return;

        int nextCycle = _cachedBPData.last_cycle_id + 1;
        int nextPlan = _cachedBPData.last_scheme_id + 1;
        int nextItem = _cachedBPData.last_item_id + 1;
        string nextYYMM = GetNextSequentialYYMM(_cachedBPData.last_recharge_alias);

        if (bpPrevCycle != null) bpPrevCycle.text = nextCycle.ToString();
        if (bpPrevPlan != null) bpPrevPlan.text = nextPlan.ToString();
        if (bpPrevItem != null) bpPrevItem.text = nextItem.ToString();
        if (bpPrevYYMM != null) bpPrevYYMM.text = nextYYMM;

        // Calculate Point Mode
        if (DateTime.TryParse(bpStartTime.value, out DateTime start) && DateTime.TryParse(bpEndTime.value, out DateTime end))
        {
            TimeSpan duration = end - start;
            int days = (int)Math.Ceiling(duration.TotalDays);
            int mode = (days >= 30) ? 900 : 800;
            int template = (mode == 900) ? _cachedBPData.last_900_cycle : _cachedBPData.last_800_cycle;
            
            if (bpPointMode != null) bpPointMode.text = $"{mode} Points ({days} days)";
            if (bpPrevTemplate != null) bpPrevTemplate.text = $"BP {template}";
        }
        else
        {
            if (bpPointMode != null) bpPointMode.text = "Enter valid dates...";
            if (bpPrevTemplate != null) bpPrevTemplate.text = "...";
        }
    }

    private string GetNextSequentialYYMM(string lastAlias)
    {
        var match = System.Text.RegularExpressions.Regex.Match(lastAlias, @"(\d{4})");
        if (!match.Success) return "2601";
        
        string yymm = match.Value;
        int yy = int.Parse(yymm.Substring(0, 2));
        int mm = int.Parse(yymm.Substring(2, 2));
        
        mm++;
        if (mm > 12) { mm = 1; yy++; }
        
        return $"{yy:D2}{mm:D2}";
    }

    private async void OnBattlePassCommitClicked()
    {
        if (_cachedBPData == null) return;

        // Validation
        if (!DateTime.TryParse(bpStartTime.value, out _) || !DateTime.TryParse(bpEndTime.value, out _))
        {
            EditorUtility.DisplayDialog("Error", "Invalid date format. Please use yyyy-mm-dd hh:mm:ss", "OK");
            return;
        }

        if (DateTime.TryParse(bpStartTime.value, out DateTime start) && DateTime.TryParse(bpEndTime.value, out DateTime end))
        {
            TimeSpan duration = end - start;
            int days = (int)Math.Ceiling(duration.TotalDays);
            int mode = (days >= 30) ? 900 : 800;
            int template = (mode == 900) ? _cachedBPData.last_900_cycle : _cachedBPData.last_800_cycle;

            var sb = new System.Text.StringBuilder();
            sb.Append("{");
            sb.Append($"\"cycle_id\":{_cachedBPData.last_cycle_id + 1},");
            sb.Append($"\"start_time\":\"{bpStartTime.value}\",");
            sb.Append($"\"end_time\":\"{bpEndTime.value}\",");
            sb.Append($"\"bg_limit\":{mode},");
            sb.Append($"\"template_cycle\":{template}");
            sb.Append("}");

            string json = sb.ToString();
            Log($"BP PAYLOAD: {json}");

            string pythonPath = "python";
            string scriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/update_bp.py");
            string tempJsonPath = Path.Combine(Path.GetTempPath(), "bp_payload.json");
            File.WriteAllText(tempJsonPath, json, System.Text.Encoding.UTF8);

            if (bpExecuteBtn != null) bpExecuteBtn.SetEnabled(false);
            if (bpExecuteBtn != null) bpExecuteBtn.text = "Committing...";

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{scriptPath}\" {GetPythonDesignPathArg()} \"{tempJsonPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(startInfo))
            {
                string output = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();
                process.WaitForExit();

                if (bpExecuteBtn != null)
                {
                    bpExecuteBtn.SetEnabled(true);
                    bpExecuteBtn.text = "Commit BattlePass Config";
                }

                if (process.ExitCode != 0)
                {
                    string errMsg = !string.IsNullOrEmpty(error) ? error : output;
                    EditorUtility.DisplayDialog("Commit Failed ❌", errMsg, "OK");
                }
                else
                {
                    EditorUtility.DisplayDialog("Commit Successful ✅", "BattlePass configuration updated across all sheets!", "Great!");
                    _cachedBPData = null; // Reset to force refresh
                    InitializeBattlePassViewAsync();
                }
            }
        }
    }

    // --- Holiday BattlePass Configurations Logic ---

    private async void InitializeHolidayBPViewAsync()
    {
        if (_cachedHBPData != null) return;

        Log("Loading Holiday BattlePass metadata...");
        
        string pythonPath = "python";
        string scriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/get_hbp_info.py");
        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            Arguments = $"\"{scriptPath}\" {GetPythonDesignPathArg()}",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (Process process = Process.Start(startInfo))
        {
            string output = await process.StandardOutput.ReadToEndAsync();
            string err = await process.StandardError.ReadToEndAsync();
            process.WaitForExit();

            if (process.ExitCode == 0)
            {
                try {
                    _cachedHBPData = JsonUtility.FromJson<HolidayBPInfo>(output);
                    UpdateHolidayBPPreview();
                    Log("Holiday BP metadata loaded successfully.");
                } catch (Exception ex) {
                    Log($"<color=red>Error parsing HBP metadata: {ex.Message}</color>");
                }
            }
            else {
                Log($"<color=red>Error fetching HBP metadata: {err}</color>");
            }
        }
    }

    private void UpdateHolidayBPPreview()
    {
        if (_cachedHBPData == null) return;

        if (hbpNextCycle != null) hbpNextCycle.text = (_cachedHBPData.last_cycle_id + 1).ToString();
        if (hbpNextRecommend != null) hbpNextRecommend.text = (_cachedHBPData.last_recommend_id + 1).ToString();
        if (hbpNextItem != null) hbpNextItem.text = (_cachedHBPData.last_item_id + 1).ToString();
        if (hbpNextScheme != null) hbpNextScheme.text = (_cachedHBPData.last_scheme_id + 1).ToString();

        if (hbpHolidayId != null && hbpSuffixPreview != null)
        {
            string hid = string.IsNullOrEmpty(hbpHolidayId.value) ? "[ID]" : hbpHolidayId.value;
            string preview = "";
            foreach (var s in _cachedHBPData.last_suffixes) {
                preview += hid + s + " ";
            }
            hbpSuffixPreview.text = preview.Trim();
        }
    }

    private async void OnHolidayBPCommitClicked()
    {
        if (string.IsNullOrEmpty(hbpHolidayId.value) || string.IsNullOrEmpty(hbpStartTime.value) || string.IsNullOrEmpty(hbpEndTime.value))
        {
            EditorUtility.DisplayDialog("Missing Input", "Please fill in Holiday ID, Start Time and End Time.", "OK");
            return;
        }

        if (hbpExecuteBtn != null) hbpExecuteBtn.SetEnabled(false);
        Log("Committing Holiday BattlePass configuration...");

        string pythonPath = "python";
        string scriptPath = Path.Combine(Application.dataPath, "Editor/EventAutomation/PythonBackend/update_hbp.py");
        var payload = new {
            holiday_id = hbpHolidayId.value,
            start_time = hbpStartTime.value,
            end_time = hbpEndTime.value
        };
        string jsonPayload = JsonUtility.ToJson(payload);
        string tempFile = Path.Combine(Path.GetTempPath(), $"hbp_payload_{Guid.NewGuid()}.json");
        File.WriteAllText(tempFile, jsonPayload);

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            Arguments = $"\"{scriptPath}\" {GetPythonDesignPathArg()} \"{tempFile}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (Process process = Process.Start(startInfo))
        {
            string output = await process.StandardOutput.ReadToEndAsync();
            string err = await process.StandardError.ReadToEndAsync();
            process.WaitForExit();

            if (hbpExecuteBtn != null) hbpExecuteBtn.SetEnabled(true);

            if (process.ExitCode == 0)
            {
                Log("<color=green>Holiday BP Commit Successful!</color>");
                EditorUtility.DisplayDialog("Success", "Holiday BattlePass configuration updated across all files.", "Great!");
                _cachedHBPData = null;
                InitializeHolidayBPViewAsync();
            }
            else
            {
                Log($"<color=red>Holiday BP Commit Failed: {err}</color>");
                EditorUtility.DisplayDialog("Commit Failed", err, "OK");
            }
        }
        
        if (File.Exists(tempFile)) File.Delete(tempFile);
    }

    // --- Settings & Path Logic ---

    private void InitializeSettingsView()
    {
        if (workspaceDropdown == null) return;

        workspaceDropdown.choices = WorkspaceOptions;
        
        string saved = EditorPrefs.GetString(WorkspacePrefKey, "WorkSpcae_Design");
        workspaceDropdown.value = saved;

        workspaceDropdown.RegisterValueChangedCallback(evt => {
            EditorPrefs.SetString(WorkspacePrefKey, evt.newValue);
            UpdateResolvedPathLabel();
            Log($"Workspace changed to: {evt.newValue}");
            
            // Invalidate caches when context changes
            _cachedHBPData = null;
            _cachedBPData = null;
            _cachedMinigamesData = null;
        });

        UpdateResolvedPathLabel();
    }

    private void UpdateResolvedPathLabel()
    {
        if (resolvedPathLabel != null)
        {
            resolvedPathLabel.text = GetActiveDesignPath();
        }
    }

    private string GetActiveDesignPath()
    {
        string workspace = EditorPrefs.GetString(WorkspacePrefKey, "WorkSpcae_Design");
        
        // Logical structure: branches/Misc/design/DesignData/[Workspace]/design/
        string root = Application.dataPath.Replace("/develop/client/Skipbo/Assets", "").Replace("\\develop\\client\\Skipbo\\Assets", "");
        string fullPath = Path.Combine(root, "design", "DesignData", workspace, "design");
        
        // Normalize slashes for windows
        return fullPath.Replace("/", "\\");
    }

    private string GetPythonDesignPathArg()
    {
        return $"\"{GetActiveDesignPath()}\"";
    }

    private string GetActiveWorkspaceName()
    {
        return EditorPrefs.GetString(WorkspacePrefKey, "WorkSpcae_Design");
    }
}
