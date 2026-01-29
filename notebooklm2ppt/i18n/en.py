TRANSLATIONS = {
    "root_title": "NotebookLM2PPT v{version} - PDF to PPT Tool",
    "startup_dialog_title": "Welcome",
    "startup_info": (
        "This software is a free and open-source PDF to PPT tool.\n\n"
        "Developer: Elliott Zheng\n\n"
        "⚠️ IMPORTANT: This software depends on the 'Smart Select' feature of [Microsoft PC Manager].\n"
        "Please ensure you have installed and started Microsoft PC Manager, and that the Smart Select feature can be activated with Ctrl+Shift+A, otherwise the conversion will not proceed.\n\n"
        "If you find this software helpful, please give it a star on GitHub or introduce it to your friends. Thank you!\n\n"
        "This software is free and open-source. If you paid for it, you have been scammed. [○･｀Д´･○]\n\n"
        "Thanks for using this tool!"
    ),
    "open_github_btn": "Open GitHub",
    "dont_show_again_btn": "Don't show again",
    "ok_btn": "OK",
    "drop_warning": "Please drop a PDF file or MinerU JSON file!",
    "file_settings_label": "📁 File Settings (Supports dragging PDF / corresponding MinerU JSON to window)",
    "pdf_file_label": "PDF File:",
    "browse_btn": "Browse...",
    "mineru_json_label": "Input MinerU JSON for PDF (Optional, for better results):",
    "info_btn": "Info",
    "output_dir_label": "Output Dir:",
    "open_btn": "Open",
    "options_label": "⚙️ Conversion Options",
    "dpi_label": "Image DPI:",
    "dpi_hint": "(150-300)",
    "delay_label": "Delay (s):",
    "delay_hint": "(After each page)",
    "timeout_label": "Timeout (s):",
    "timeout_hint": "(Max per page)",
    "ratio_label": "Window Scale:",
    "ratio_hint": "(Suggest 0.7-0.9)",
    "inpaint_label": "Remove Watermark",
    "inpaint_method_label": "Inpaint Method:",
    "image_only_label": "Image Only Mode (Skip smart selection, PPT non-editable)",
    "force_regenerate_label": "Force Regenerate All Pages",
    "unify_font_label": "Unify Font",
    "font_name_label": "Target Font:",
    "page_range_label": "Page Range:",
    "page_range_hint": "Empty=All, e.g., 1-3,5,7-9",
    "button_offset_label": "Btn Offset (px):",
    "calibrate_label": "Calibrate Button Position",
    "core_param_warning": "⚠️ Core: The program simulates mouse clicks on 'Convert to PPT' button.",
    "core_param_warning2": "   If button position is incorrect, it won't work! Use 'Calibrate' to fix.",
    "core_param_warning3": "   Tip: Calibration results are saved automatically.",
    "start_btn": "🚀 Start Queue",
    "stop_btn": "⏹️ Stop Queue",
    "log_area_label": "📋 Run Log",
    "select_pdf_title": "Select PDF File",
    "select_json_title": "Select MinerU JSON File",
    "select_output_title": "Select Output Directory",
    "set_new_dir_msg": "New directory set: {directory}",
    "set_output_dir_warning": "Please set output directory first",
    "create_output_dir_msg": "Output directory created: {output_dir}",
    "create_output_dir_error": "Failed to create output directory: {error}",
    "open_output_dir_error": "Failed to open output directory: {error}",
    "image_only_confirm_title": "Confirm Image Only Mode",
    "image_only_confirm_msg": (
        "Image Only Mode will:\n\n"
        "• Skip smart selection feature\n"
        "• Directly insert watermarked PNG images into PPT\n"
        "• Not generate editable text content\n"
        "• Be faster, but PPT content will be non-editable\n\n"
        "Continue with Image Only Mode?"
    ),
    "inpaint_method_info_title": "Inpaint Method Description",
    "inpaint_method_info_prefix": "Inpaint Method Description:\n",
    "close_btn": "Close",
    "select_pdf_error": "Please select a PDF file first",
    "stopping_msg": "Stopping conversion...",
    "config_saved": "Configuration saved to disk",
    "config_save_fail": "Failed to save configuration: {error}",
    "config_load_fail": "Failed to load configuration: {error}",
    "default_config_created": "Default configuration file created",
    "saved_offset": "Saved: {offset}px",
    "unsaved_offset": "Unsaved: Will auto-calibrate",
    "mineru_info_title": "About MinerU",
    "mineru_info_content": (
        "MinerU is an online document parsing tool.\n\n"
        "Steps:\n"
        "1. Upload your PDF to MinerU website https://mineru.net/ and wait for parsing.\n"
        "2. Download the generated JSON file.\n"
        "3. Select the JSON file in 'Input MinerU JSON for PDF (Optional)'.\n\n"
        "Description: This JSON contains info like page structure, text, and layout; this program uses it to further optimize PPT images, backgrounds, and text.\n\n"
        "Note: Ensure the JSON corresponds to the PDF, otherwise optimization may be incorrect."
    ),
    "open_mineru_website": "Open MinerU Website",
    "start_processing": "Starting process: {file}",
    "page_range_error": "Page range format error, please use formats like 1-3,5,7-",
    "image_only_mode_start": "Image Only Mode: Directly inserting PNG images into PPT",
    "conversion_stopped_msg": "Conversion stopped by user",
    "conversion_stopped_title": "Conversion Stopped",
    "mineru_optimizing": "Optimizing PPT with MinerU info: {file}",
    "refine_ppt_done": "✅ refine_ppt completed",
    "refine_extra_msg": "Original PPT saved in the same directory",
    "conversion_done": "✅ Conversion completed!",
    "output_file": "📄 Output file: {file}",
    "conversion_success_title": "Conversion Success",
    "conversion_success_msg": "PDF successfully converted to PPT!\n\nFile location:\n{file}",
    "conversion_fail": "❌ Conversion failed: {error}",
    "conversion_fail_title": "Conversion Failed",
    "conversion_fail_msg": "An error occurred during processing:\n{error}",
    "integer_offset_error": "Button offset must be an integer or empty",
    "language_menu": "Language",
    "lang_zh_cn": "简体中文",
    "lang_en": "English",
    "cut": "Cut",
    "copy": "Copy",
    "paste": "Paste",
    "select_all": "Select All",
    "error_btn": "Error",
    "file_added_msg": "File added: {file}",
    "drag_drop_warning": "Please drag and drop PDF or Mineru JSON files!",
    "offset_value_error": "Done button offset must be an integer or empty",
    "method_background_smooth_name": "Smart Smooth (Recommended)",
    "method_background_smooth_desc": "Best overall effect, suitable for most text and watermark removal scenarios",
    "method_edge_mean_smooth_name": "Edge Mean Fill",
    "method_edge_mean_smooth_desc": "Fills with average color of surrounding pixels, suitable for solid or simple backgrounds",
    "method_background_name": "Fast Solid Fill",
    "method_background_desc": "Directly fills with a single background color, only suitable for simple backgrounds, fastest speed",
    "method_onion_name": "Onion Skin Repair",
    "method_onion_desc": "Repairs layer by layer from outside in, suitable for thin scratches or lines",
    "method_griddata_name": "Gradient Interpolation",
    "method_griddata_desc": "Calculates smooth surface transitions, suitable for backgrounds with gradients",
    "method_skimage_name": "Biharmonic Repair",
    "method_skimage_desc": "Computationally intensive and slow, but better at maintaining lighting continuity",
    "calibration_dialog_title": "Tip",
    "calibration_dialog_msg": (
        "Button position calibration in progress. A confirmation is required. Please read this carefully.\n\n"
        "⚠️ Please ensure that [Microsoft PC Manager] is installed and running, and that the 'Smart Select' feature can be activated via Ctrl+Shift+A. Otherwise, the Smart Select toolbar will not appear!\n\n"
        'After clicking "OK", do not move the mouse. Wait for the Smart Select toolbar to appear, then manually move the mouse and click the "Smart Copy to PPT" button to complete the calibration.\n\n'
        'Click "OK" when ready. Do not minimize the window or interfere with mouse operations during the conversion.'
    ),
    "pc_manager_not_running_msg": (
        "Microsoft PC Manager is not running.\n\n"
        "This tool depends on the 'Smart Select / Smart Copy to PPT' feature of "
        "Microsoft PC Manager to automatically capture pages and generate PPT. "
        "If PC Manager is not installed or not running, the automatic conversion "
        "cannot work.\n\n"
        "Please install and start Microsoft PC Manager (process name: MSPCManager.exe), "
        "then try the conversion again."
    ),
    "pc_manager_open_website_confirm": "Open the Microsoft PC Manager official website to download it?",
    "open_pc_manager_website_error": "Failed to open Microsoft PC Manager website: {error}",
    "queue_label": "🗂 Batch Queue",
    "queue_col_id": "ID",
    "queue_col_pdf": "PDF",
    "queue_col_json": "JSON",
    "queue_col_status": "Status",
    "queue_col_output": "Output",
    "queue_add_task": "Add Task",
    "queue_add_multi_pdf": "Add Multiple PDFs",
    "queue_remove_selected": "Remove Selected",
    "queue_clear": "Clear Queue",
    "queue_start": "Start Queue",
    "queue_stop": "Stop Queue",
    "queue_status_pending": "Pending",
    "queue_status_running": "Running",
    "queue_status_done": "Done",
    "queue_status_error": "Error",
    "queue_task_added": "Task added: {file}",
    "queue_task_done": "Task done, output: {file}",
    "task_details_title": "Task Details",
    "none": "None",
    "queue_task_updated": "Task updated (Duplicate PDF): {file}",
    "queue_task_removed": "Selected tasks removed",
    "queue_cleared": "Queue cleared",
    "queue_started": "Queue started",
    "queue_stopping": "Stopping queue...",
    "queue_stopped": "Queue stopped",
    "queue_finished": "Queue finished",
    "queue_empty_msg": "Queue is empty! Please add tasks first, or select a PDF above and click 'Start Queue'.",
    "task_settings_title": "⚙️ Task Conversion Settings",
    "yes": "Yes",
    "no": "No",
    "save_btn": "Save Changes",
    "global_settings_label": "🌐 Global Settings",
    "automation_settings_label": "🤖 Automation Settings (Applies to all tasks)",
    "automation_warning": "These settings apply to all tasks, please ensure PC Manager is running",
    "add_task_title": "Add New Task",
    "task_params_label": "⚙️ Task Conversion Parameters",
    "add_btn": "Add Task",
    "cancel_btn": "Cancel",
    "drag_drop_added": "Added {count} task(s) to queue",
}
