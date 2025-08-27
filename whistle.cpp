#include "whistle.h"

// XMLSpreadsheetWriter implementation
std::string XMLSpreadsheetWriter::escapeXML(const std::string& text) {
    if (text.empty()) {
        return text;
    }
    
    std::string escaped;
    escaped.reserve(text.length() * 2); // Reserve extra space for escapes
    
    for (size_t i = 0; i < text.length(); ++i) {
        char c = text[i];
        switch (c) {
            case '&':
                escaped += "&amp;";
                break;
            case '<':
                escaped += "&lt;";
                break;
            case '>':
                escaped += "&gt;";
                break;
            case '"':
                escaped += "&quot;";
                break;
            case '\'':
                escaped += "&apos;";
                break;
            default:
                escaped += c;
                break;
        }
    }
    
    return escaped;
}

std::string XMLSpreadsheetWriter::cleanSheetName(const std::string& name) {
    if (name.empty()) {
        return "Sheet1";
    }
    
    std::string clean = name;
    
    // Replace invalid characters with bounds checking
    for (size_t i = 0; i < clean.length(); ++i) {
        char& c = clean[i];
        if (c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']' || c == ':') {
            c = '_';
        }
    }
    
    // Limit to 31 characters (Excel limit)
    if (clean.length() > 31) {
        clean = clean.substr(0, 31);
    }
    
    return clean;
}

XMLSpreadsheetWriter::XMLSpreadsheetWriter(const std::string& filename) : file(filename) {}

XMLSpreadsheetWriter::~XMLSpreadsheetWriter() {
    if (file.is_open()) {
        file.close();
    }
}

void XMLSpreadsheetWriter::addWorksheet(const std::string& name) {
    std::string clean_name = cleanSheetName(name);
    worksheets[clean_name] = std::vector<std::vector<std::string>>();
}

void XMLSpreadsheetWriter::addRow(const std::string& worksheet_name, const std::vector<std::string>& row) {
    std::string clean_name = cleanSheetName(worksheet_name);
    if (worksheets.find(clean_name) != worksheets.end()) {
        worksheets[clean_name].push_back(row);
    }
}

bool XMLSpreadsheetWriter::writeFile() {
    if (!file.is_open()) {
        return false;
    }
    
    // Write XML header
    file << "<?xml version=\"1.0\"?>" << std::endl;
    file << "<?mso-application progid=\"Excel.Sheet\"?>" << std::endl;
    file << "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"" << std::endl;
    file << " xmlns:o=\"urn:schemas-microsoft-com:office:office\"" << std::endl;
    file << " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"" << std::endl;
    file << " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"" << std::endl;
    file << " xmlns:html=\"http://www.w3.org/TR/REC-html40\">" << std::endl;
    
    // Write document properties
    file << " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">" << std::endl;
    file << "  <Created>" << std::chrono::system_clock::now().time_since_epoch().count() << "</Created>" << std::endl;
    file << "  <Application>Regex Analyzer</Application>" << std::endl;
    file << " </DocumentProperties>" << std::endl;
    
    // Write styles
    file << " <Styles>" << std::endl;
    file << "  <Style ss:ID=\"Header\">" << std::endl;
    file << "   <Font ss:Bold=\"1\"/>" << std::endl;
    file << "   <Interior ss:Color=\"#C0C0C0\" ss:Pattern=\"Solid\"/>" << std::endl;
    file << "   <Borders>" << std::endl;
    file << "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "   </Borders>" << std::endl;
    file << "  </Style>" << std::endl;
    file << "  <Style ss:ID=\"Cell\">" << std::endl;
    file << "   <Borders>" << std::endl;
    file << "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>" << std::endl;
    file << "   </Borders>" << std::endl;
    file << "   <Alignment ss:Vertical=\"Top\" ss:WrapText=\"1\"/>" << std::endl;
    file << "  </Style>" << std::endl;
    file << " </Styles>" << std::endl;
    
    // Write worksheets
    for (const auto& [sheet_name, rows] : worksheets) {
        file << " <Worksheet ss:Name=\"" << escapeXML(sheet_name) << "\">" << std::endl;
        file << "  <Table>" << std::endl;
        
        // Set column widths
        file << "   <Column ss:Width=\"120\"/>" << std::endl; // Finding
        file << "   <Column ss:Width=\"240\"/>" << std::endl; // File
        file << "   <Column ss:Width=\"60\"/>" << std::endl;  // Line
        file << "   <Column ss:Width=\"120\"/>" << std::endl; // Comments
        file << "   <Column ss:Width=\"90\"/>" << std::endl;  // Ease
        file << "   <Column ss:Width=\"90\"/>" << std::endl;  // Significance
        file << "   <Column ss:Width=\"90\"/>" << std::endl;  // Risk
        file << "   <Column ss:Width=\"360\"/>" << std::endl; // Statement
        
        for (size_t i = 0; i < rows.size(); ++i) {
            const auto& row = rows[i];
            file << "   <Row>" << std::endl;
            
            for (size_t j = 0; j < row.size(); ++j) {
                std::string style_id = (i == 0) ? "Header" : "Cell";
                std::string cell_data = escapeXML(row[j]);
                
                // Check if it's a number (for line numbers)
                bool is_number = false;
                if (j == 2 && i > 0) { // Line number column, not header
                    try {
                        std::stoi(row[j]);
                        is_number = true;
                    } catch (...) {
                        is_number = false;
                    }
                }
                
                file << "    <Cell ss:StyleID=\"" << style_id << "\">" << std::endl;
                if (is_number) {
                    file << "     <Data ss:Type=\"Number\">" << cell_data << "</Data>" << std::endl;
                } else {
                    file << "     <Data ss:Type=\"String\">" << cell_data << "</Data>" << std::endl;
                }
                file << "    </Cell>" << std::endl;
            }
            
            file << "   </Row>" << std::endl;
        }
        
        file << "  </Table>" << std::endl;
        
        // Add worksheet options (freeze header row so it stays visible when scrolling)
        if (!rows.empty()) {
            file << "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">" << std::endl;
            file << "   <FreezePanes/>" << std::endl;
            file << "   <FrozenNoSplit/>" << std::endl;
            file << "   <SplitHorizontal>1</SplitHorizontal>" << std::endl;
            file << "   <TopRowBottomPane>1</TopRowBottomPane>" << std::endl;
            file << "   <ActivePane>2</ActivePane>" << std::endl;
            file << "  </WorksheetOptions>" << std::endl;
        }
        
        file << " </Worksheet>" << std::endl;
    }
    
    file << "</Workbook>" << std::endl;
    
    return true;
}

bool XMLSpreadsheetWriter::isOpen() const {
    return file.is_open();
}

// ProgressTracker implementation
void ProgressTracker::setTotal(int t) {
    total = t;
    start_time = std::chrono::steady_clock::now();
}

void ProgressTracker::increment() {
    processed++;
    printProgress();
}

void ProgressTracker::printProgress() const {
    std::lock_guard<std::mutex> lock(print_mutex);
    
    int proc = processed.load();
    int tot = total.load();
    
    if (tot == 0) return;
    
    auto now = std::chrono::steady_clock::now();
    auto elapsed = std::chrono::duration_cast<std::chrono::seconds>(now - start_time).count();
    
    double percentage = (double)proc / tot * 100.0;
    int remaining = tot - proc;
    
    // Estimate time remaining
    double eta_seconds = 0;
    if (proc > 0 && elapsed > 0) {
        double rate = (double)proc / elapsed;
        eta_seconds = remaining / rate;
    }
    
    std::cout << "\r[" << std::setw(3) << std::fixed << std::setprecision(1) 
              << percentage << "%] Processed: " << proc << "/" << tot 
              << " | Remaining: " << remaining;
    
    if (eta_seconds > 0) {
        int eta_min = (int)(eta_seconds / 60);
        int eta_sec = (int)(eta_seconds) % 60;
        std::cout << " | ETA: " << eta_min << "m " << eta_sec << "s";
    }
    
    std::cout << std::flush;
    
    if (proc == tot) {
        std::cout << std::endl << "Processing complete!" << std::endl;
    }
}

// RegexAnalyzer implementation
std::vector<ExpressionPattern> RegexAnalyzer::loadExpressions(const std::string& filename) {
    std::vector<ExpressionPattern> patterns;
    std::ifstream file(filename);
    
    if (!file.is_open()) {
        throw std::runtime_error("Could not open expressions.properties file");
    }
    
    std::string line;
    bool in_expressions_section = false;
    
    while (std::getline(file, line)) {
        // Trim whitespace
        line.erase(0, line.find_first_not_of(" \t\r\n"));
        line.erase(line.find_last_not_of(" \t\r\n") + 1);
        
        // Skip empty lines and comments
        if (line.empty() || line[0] == '#') continue;
        
        // Check for [expressions] section
        if (line == "[expressions]") {
            in_expressions_section = true;
            continue;
        }
        
        // Check for other sections
        if (line[0] == '[' && line.back() == ']') {
            in_expressions_section = false;
            continue;
        }
        
        if (in_expressions_section) {
            // Parse expression.name=pattern
            size_t eq_pos = line.find('=');
            if (eq_pos != std::string::npos) {
                std::string key = line.substr(0, eq_pos);
                std::string value = line.substr(eq_pos + 1);
                
                // Trim key and value
                key.erase(0, key.find_first_not_of(" \t"));
                key.erase(key.find_last_not_of(" \t") + 1);
                value.erase(0, value.find_first_not_of(" \t"));
                value.erase(value.find_last_not_of(" \t") + 1);
                
                if (key.substr(0, 11) == "expression.") {
                    std::string expr_name = key.substr(11);
                    try {
                        // Check for inline flags like (?i) at the beginning
                        std::regex_constants::syntax_option_type flags = std::regex_constants::ECMAScript;
                        std::string pattern_str = value;
                        
                        // Handle (?i) case-insensitive flag
                        if (pattern_str.substr(0, 4) == "(?i)") {
                            flags |= std::regex_constants::icase;
                            pattern_str = pattern_str.substr(4); // Remove (?i) from pattern
                        }
                        // Handle (?-i) case-sensitive flag (explicit)
                        else if (pattern_str.substr(0, 5) == "(?-i)") {
                            // Default is case-sensitive, so just remove the flag
                            pattern_str = pattern_str.substr(5); // Remove (?-i) from pattern
                        }
                        // Default case-insensitive behavior (as originally implemented)
                        else {
                            flags |= std::regex_constants::icase;
                        }
                        
                        std::regex pattern(pattern_str, flags);
                        patterns.push_back({expr_name, std::move(pattern)});
                        std::cout << "Loaded expression: " << expr_name << " = " << value << std::endl;
                    } catch (const std::regex_error& e) {
                        std::cerr << "Invalid regex for " << expr_name << ": " << value 
                                 << " Error: " << e.what() << std::endl;
                    }
                }
            }
        }
    }
    
    return patterns;
}

bool RegexAnalyzer::isTextFile(const std::string& filepath) {
    std::ifstream file(filepath, std::ios::binary);
    if (!file.is_open()) {
        return false;
    }
    
    // Read first chunk of file to analyze
    const size_t sample_size = 8192; // 8KB sample
    char buffer[8192];
    
    file.read(buffer, sample_size);
    std::streamsize bytes_read = file.gcount();
    
    if (bytes_read <= 0) {
        return false; // Empty file or read error
    }
    
    // Check for null bytes (common in binary files)
    int null_count = 0;
    int printable_count = 0;
    int control_count = 0;
    
    // Only iterate over the actual bytes read with bounds check
    for (std::streamsize i = 0; i < bytes_read; ++i) {
        unsigned char byte = static_cast<unsigned char>(buffer[i]);
        
        if (byte == 0) {
            null_count++;
        } else if (std::isprint(byte) || byte == '\t' || byte == '\n' || byte == '\r') {
            printable_count++;
        } else if (byte < 32 || byte == 127) {
            control_count++;
        }
    }
    
    // Heuristic: if more than 5% null bytes, likely binary
    if (null_count > bytes_read * 0.05) {
        return false;
    }
    
    // Heuristic: if less than 70% printable characters, likely binary
    double printable_ratio = static_cast<double>(printable_count) / bytes_read;
    if (printable_ratio < 0.70) {
        return false;
    }
    
    // Additional check for UTF-8 BOM - ensure we have enough bytes
    if (bytes_read >= 3 && 
        static_cast<unsigned char>(buffer[0]) == 0xEF &&
        static_cast<unsigned char>(buffer[1]) == 0xBB &&
        static_cast<unsigned char>(buffer[2]) == 0xBF) {
        return true; // UTF-8 BOM indicates text file
    }
    
    return true; // Passed all heuristics
}

void RegexAnalyzer::processFile(const std::string& filepath) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        return;
    }
    
    std::string line;
    line.reserve(1024); // Pre-allocate reasonable size
    int line_number = 0;
    std::vector<Finding> local_findings;
    local_findings.reserve(100); // Pre-allocate to reduce reallocations
    
    // Limit line length to prevent excessive memory usage
    const size_t MAX_LINE_LENGTH = 100000; // 100KB max per line
    
    while (std::getline(file, line)) {
        line_number++;
        
        // Skip extremely long lines that might cause issues
        if (line.length() > MAX_LINE_LENGTH) {
            std::cerr << "Warning: Skipping very long line " << line_number 
                     << " in file " << filepath << " (length: " << line.length() << ")" << std::endl;
            continue;
        }
        
        // Process each expression safely
        for (size_t expr_idx = 0; expr_idx < expressions.size(); ++expr_idx) {
            const auto& expr = expressions[expr_idx];
            try {
                std::smatch match;
                if (std::regex_search(line, match, expr.pattern)) {
                    Finding finding;
                    finding.expression_name = expr.name;
                    finding.filename = filepath;
                    finding.line_number = line_number;
                    finding.actual_match = match.str();  // The actual matched text
                    finding.statement = line;            // The full line
                    local_findings.push_back(std::move(finding));
                }
            } catch (const std::exception& e) {
                std::cerr << "Error processing regex '" << expr.name 
                         << "' on line " << line_number 
                         << " in file " << filepath << ": " << e.what() << std::endl;
                continue; // Skip this regex and continue with others
            }
        }
    }
    
    if (!local_findings.empty()) {
        std::lock_guard<std::mutex> lock(findings_mutex);
        all_findings.insert(all_findings.end(), 
                           std::make_move_iterator(local_findings.begin()),
                           std::make_move_iterator(local_findings.end()));
    }
    
    progress.increment();
}

std::vector<std::string> RegexAnalyzer::findTextFiles(const std::string& directory) {
    std::vector<std::string> text_files;
    
    try {
        if (!std::filesystem::exists(directory)) {
            std::cerr << "Error: Directory does not exist: " << directory << std::endl;
            return text_files;
        }
        
        if (!std::filesystem::is_directory(directory)) {
            std::cerr << "Error: Path is not a directory: " << directory << std::endl;
            return text_files;
        }
        
        for (const auto& entry : std::filesystem::recursive_directory_iterator(directory)) {
            try {
                if (entry.is_regular_file()) {
                    if (isTextFile(entry.path().string())) {
                        text_files.push_back(entry.path().string());
                    }
                }
            } catch (const std::filesystem::filesystem_error& e) {
                std::cerr << "Error accessing file: " << entry.path() 
                         << " - " << e.what() << std::endl;
                continue; // Skip this file and continue
            }
        }
    } catch (const std::filesystem::filesystem_error& e) {
        std::cerr << "Error accessing directory: " << e.what() << std::endl;
    }
    
    return text_files;
}

void RegexAnalyzer::workerThread() {
    while (true) {
        std::string filepath;
        
        {
            std::lock_guard<std::mutex> lock(queue_mutex);
            if (file_queue.empty()) {
                break;
            }
            filepath = file_queue.back();
            file_queue.pop_back();
        }
        
        processFile(filepath);
    }
}

void RegexAnalyzer::analyze(const std::string& directory, const std::string& expressions_file, 
            const std::string& output_file, int num_threads) {
    
    std::cout << "Loading expressions from: " << expressions_file << std::endl;
    expressions = loadExpressions(expressions_file);
    
    if (expressions.empty()) {
        throw std::runtime_error("No valid expressions found in properties file");
    }
    
    std::cout << "Loaded " << expressions.size() << " expressions" << std::endl;
    std::cout << "Scanning directory: " << directory << std::endl;
    
    file_queue = findTextFiles(directory);
    std::cout << "Found " << file_queue.size() << " text files" << std::endl;
    
    if (file_queue.empty()) {
        std::cout << "No text files found to process" << std::endl;
        return;
    }
    
    progress.setTotal(file_queue.size());
    std::cout << "Starting analysis with " << num_threads << " threads..." << std::endl;
    
    // Launch worker threads
    std::vector<std::thread> threads;
    for (int i = 0; i < num_threads; ++i) {
        threads.emplace_back(&RegexAnalyzer::workerThread, this);
    }
    
    // Wait for all threads to complete
    for (auto& thread : threads) {
        thread.join();
    }
    
    std::cout << std::endl << "Analysis complete. Found " << all_findings.size() << " matches" << std::endl;
    std::cout << "Writing results to: " << output_file << std::endl;
    
    writeResults(output_file);
}

void RegexAnalyzer::writeResults(const std::string& output_filename) {
#if USE_XLSX
    writeXLSXResults(output_filename);
#else
    writeXMLSpreadsheetResults(output_filename);
#endif
}

#if USE_XLSX
void RegexAnalyzer::writeXLSXResults(const std::string& output_filename) {
    // Create workbook
    lxw_workbook* workbook = workbook_new(output_filename.c_str());
    if (!workbook) {
        throw std::runtime_error("Failed to create Excel workbook: " + output_filename);
    }
    
    // Create formats
    lxw_format* header_format = workbook_add_format(workbook);
    format_set_bold(header_format);
    format_set_bg_color(header_format, LXW_COLOR_GRAY);
    format_set_border(header_format, LXW_BORDER_THIN);
    
    lxw_format* cell_format = workbook_add_format(workbook);
    format_set_border(cell_format, LXW_BORDER_THIN);
    format_set_text_wrap(cell_format);
    
    // Group findings by expression
    std::map<std::string, std::vector<Finding>> grouped_findings;
    
    for (const auto& finding : all_findings) {
        grouped_findings[finding.expression_name].push_back(finding);
    }
    
    // Create worksheet for each expression
    for (const auto& [expr_name, findings] : grouped_findings) {
        // Clean sheet name (Excel has restrictions on sheet names)
        std::string sheet_name = expr_name;
        // Replace invalid characters
        for (char& c : sheet_name) {
            if (c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']' || c == ':') {
                c = '_';
            }
        }
        // Limit to 31 characters (Excel limit)
        if (sheet_name.length() > 31) {
            sheet_name = sheet_name.substr(0, 31);
        }
        
        lxw_worksheet* worksheet = workbook_add_worksheet(workbook, sheet_name.c_str());
        if (!worksheet) {
            std::cerr << "Failed to create worksheet: " << sheet_name << std::endl;
            continue;
        }
        
        // Set column widths
        worksheet_set_column(worksheet, 0, 0, 20, nullptr); // Finding
        worksheet_set_column(worksheet, 1, 1, 40, nullptr); // File
        worksheet_set_column(worksheet, 2, 2, 10, nullptr); // Line
        worksheet_set_column(worksheet, 3, 3, 20, nullptr); // Comments
        worksheet_set_column(worksheet, 4, 4, 15, nullptr); // Ease
        worksheet_set_column(worksheet, 5, 5, 15, nullptr); // Significance
        worksheet_set_column(worksheet, 6, 6, 15, nullptr); // Risk
        worksheet_set_column(worksheet, 7, 7, 60, nullptr); // Statement
        
        // Write headers
        worksheet_write_string(worksheet, 0, 0, "Finding", header_format);
        worksheet_write_string(worksheet, 0, 1, "File", header_format);
        worksheet_write_string(worksheet, 0, 2, "Line", header_format);
        worksheet_write_string(worksheet, 0, 3, "Comments", header_format);
        worksheet_write_string(worksheet, 0, 4, "Ease", header_format);
        worksheet_write_string(worksheet, 0, 5, "Significance", header_format);
        worksheet_write_string(worksheet, 0, 6, "Risk", header_format);
        worksheet_write_string(worksheet, 0, 7, "Statement", header_format);
        
        // Write findings
        int row = 1;
        for (const auto& finding : findings) {
            worksheet_write_string(worksheet, row, 0, finding.actual_match.c_str(), cell_format);  // Actual match
            worksheet_write_string(worksheet, row, 1, finding.filename.c_str(), cell_format);
            worksheet_write_number(worksheet, row, 2, finding.line_number, cell_format);
            worksheet_write_string(worksheet, row, 3, "", cell_format); // Comments (blank)
            worksheet_write_string(worksheet, row, 4, "", cell_format); // Ease (blank)
            worksheet_write_string(worksheet, row, 5, "", cell_format); // Significance (blank)
            worksheet_write_string(worksheet, row, 6, "", cell_format); // Risk (blank)
            worksheet_write_string(worksheet, row, 7, finding.statement.c_str(), cell_format);     // Full line
            row++;
        }
        
        // Freeze the header row so it stays visible when scrolling
        worksheet_freeze_panes(worksheet, 1, 0);
        
        std::cout << "Created sheet: " << sheet_name << " with " << findings.size() << " findings" << std::endl;
    }
    
    // Create a summary worksheet with all findings
    if (!all_findings.empty()) {
        lxw_worksheet* summary_worksheet = workbook_add_worksheet(workbook, "Summary");
        if (summary_worksheet) {
            // Set column widths
            worksheet_set_column(summary_worksheet, 0, 0, 20, nullptr); // Finding
            worksheet_set_column(summary_worksheet, 1, 1, 40, nullptr); // File
            worksheet_set_column(summary_worksheet, 2, 2, 10, nullptr); // Line
            worksheet_set_column(summary_worksheet, 3, 3, 20, nullptr); // Comments
            worksheet_set_column(summary_worksheet, 4, 4, 15, nullptr); // Ease
            worksheet_set_column(summary_worksheet, 5, 5, 15, nullptr); // Significance
            worksheet_set_column(summary_worksheet, 6, 6, 15, nullptr); // Risk
            worksheet_set_column(summary_worksheet, 7, 7, 60, nullptr); // Statement
            
            // Write headers
            worksheet_write_string(summary_worksheet, 0, 0, "Finding", header_format);
            worksheet_write_string(summary_worksheet, 0, 1, "File", header_format);
            worksheet_write_string(summary_worksheet, 0, 2, "Line", header_format);
            worksheet_write_string(summary_worksheet, 0, 3, "Comments", header_format);
            worksheet_write_string(summary_worksheet, 0, 4, "Ease", header_format);
            worksheet_write_string(summary_worksheet, 0, 5, "Significance", header_format);
            worksheet_write_string(summary_worksheet, 0, 6, "Risk", header_format);
            worksheet_write_string(summary_worksheet, 0, 7, "Statement", header_format);
            
            // Write all findings
            int row = 1;
            for (const auto& finding : all_findings) {
                worksheet_write_string(summary_worksheet, row, 0, finding.actual_match.c_str(), cell_format);  // Actual match
                worksheet_write_string(summary_worksheet, row, 1, finding.filename.c_str(), cell_format);
                worksheet_write_number(summary_worksheet, row, 2, finding.line_number, cell_format);
                worksheet_write_string(summary_worksheet, row, 3, "", cell_format); // Comments (blank)
                worksheet_write_string(summary_worksheet, row, 4, "", cell_format); // Ease (blank)
                worksheet_write_string(summary_worksheet, row, 5, "", cell_format); // Significance (blank)
                worksheet_write_string(summary_worksheet, row, 6, "", cell_format); // Risk (blank)
                worksheet_write_string(summary_worksheet, row, 7, finding.statement.c_str(), cell_format);     // Full line
                row++;
            }
            
            // Freeze the header row so it stays visible when scrolling
            worksheet_freeze_panes(summary_worksheet, 1, 0);
            
            std::cout << "Created Summary sheet with " << all_findings.size() << " total findings" << std::endl;
        }
    }
    
    // Close workbook
    lxw_error error = workbook_close(workbook);
    if (error != LXW_NO_ERROR) {
        throw std::runtime_error("Failed to save Excel workbook: " + std::string(lxw_strerror(error)));
    }
    
    std::cout << "Successfully created Excel file: " << output_filename << std::endl;
}
#endif

void RegexAnalyzer::writeXMLSpreadsheetResults(const std::string& output_filename) {
    // Ensure .xml extension
    std::string xml_filename = output_filename;
    if (xml_filename.find_last_of('.') == std::string::npos) {
        xml_filename += ".xml";
    } else {
        size_t dot_pos = xml_filename.find_last_of('.');
        std::string ext = xml_filename.substr(dot_pos);
        if (ext != ".xml") {
            xml_filename = xml_filename.substr(0, dot_pos) + ".xml";
        }
    }
    
    XMLSpreadsheetWriter writer(xml_filename);
    if (!writer.isOpen()) {
        throw std::runtime_error("Failed to create XML spreadsheet: " + xml_filename);
    }
    
    // Group findings by expression
    std::map<std::string, std::vector<Finding>> grouped_findings;
    
    for (const auto& finding : all_findings) {
        grouped_findings[finding.expression_name].push_back(finding);
    }
    
    // Create worksheet for each expression
    for (const auto& [expr_name, findings] : grouped_findings) {
        writer.addWorksheet(expr_name);
        
        // Add header row
        writer.addRow(expr_name, {"Finding", "File", "Line", "Comments", "Ease", "Significance", "Risk", "Statement"});
        
        // Add findings
        for (const auto& finding : findings) {
            writer.addRow(expr_name, {
                finding.actual_match,        // Actual regex match
                finding.filename,
                std::to_string(finding.line_number),
                "", // Comments (blank)
                "", // Ease (blank)
                "", // Significance (blank)
                "", // Risk (blank)
                finding.statement            // Full line
            });
        }
        
        std::cout << "Created sheet: " << expr_name << " with " << findings.size() << " findings" << std::endl;
    }
    
    // Create a summary worksheet with all findings
    if (!all_findings.empty()) {
        writer.addWorksheet("Summary");
        
        // Add header row
        writer.addRow("Summary", {"Finding", "File", "Line", "Comments", "Ease", "Significance", "Risk", "Statement"});
        
        // Add all findings
        for (const auto& finding : all_findings) {
            writer.addRow("Summary", {
                finding.actual_match,        // Actual regex match
                finding.filename,
                std::to_string(finding.line_number),
                "", // Comments (blank)
                "", // Ease (blank)
                "", // Significance (blank)
                "", // Risk (blank)
                finding.statement            // Full line
            });
        }
        
        std::cout << "Created Summary sheet with " << all_findings.size() << " total findings" << std::endl;
    }
    
    if (!writer.writeFile()) {
        throw std::runtime_error("Failed to write XML spreadsheet file");
    }
    
    std::cout << "Successfully created XML Spreadsheet file: " << xml_filename << std::endl;
    std::cout << "This file can be opened in Excel, LibreOffice Calc, or Google Sheets" << std::endl;
}

void printUsage(const char* program_name) {
    std::cout << "Usage: " << program_name << " <directory> <expressions_file> <output_file> [num_threads]" << std::endl;
    std::cout << "  directory:        Directory to search for text files" << std::endl;
    std::cout << "  expressions_file: Path to expressions.properties file" << std::endl;
    std::cout << "  output_file:      Base name for output files" << std::endl;
    std::cout << "  num_threads:      Number of worker threads (default: 4)" << std::endl;
    std::cout << std::endl;
    std::cout << "Example expressions.properties format:" << std::endl;
    std::cout << "[expressions]" << std::endl;
    std::cout << "expression.url=https?://[\\w.-]+[\\w/]+" << std::endl;
    std::cout << "expression.ip=\\b(?:[0-9]{1,3}\\.){3}[0-9]{1,3}\\b" << std::endl;
}

int main(int argc, char* argv[]) {
    if (argc < 4 || argc > 5) {
        printUsage(argv[0]);
        return 1;
    }
    
    std::string directory = argv[1];
    std::string expressions_file = argv[2];
    std::string output_file = argv[3];
    int num_threads = (argc == 5) ? std::stoi(argv[4]) : 4;
    
#if USE_XLSX
    std::cout << "Using XLSX output format" << std::endl;
#else
    std::cout << "Using XML Spreadsheet 2003 output format (XLSX library not available)" << std::endl;
#endif
    
    try {
        RegexAnalyzer analyzer;
        analyzer.analyze(directory, expressions_file, output_file, num_threads);
        
        std::cout << "Analysis completed successfully!" << std::endl;
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
}

// Compilation instructions:
// For systems with libxlsxwriter:
//   g++ -std=c++17 -pthread -O2 -DHAVE_XLSXWRITER -o whistle whistle.cpp -lxlsxwriter
//
// For systems without libxlsxwriter (RHEL8, etc) - uses XML Spreadsheet 2003:
//   g++ -std=c++17 -pthread -O2 -o whistle whistle.cpp
//
// Build from source libxlsxwriter on RHEL8:
//   wget https://github.com/jmcnamara/libxlsxwriter/archive/RELEASE_1.1.5.tar.gz
//   tar -xzf RELEASE_1.1.5.tar.gz
//   cd libxlsxwriter-RELEASE_1.1.5
//   make
//   sudo make install
//   sudo ldconfig
//   g++ -std=c++17 -pthread -O2 -DHAVE_XLSXWRITER -o whistle whistle.cpp -lxlsxwriter
