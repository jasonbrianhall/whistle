#include "whistle.h"

// DebugLogger implementation
DebugLogger::DebugLogger(const std::string& filename) : enabled(true) {
    log_file.open(filename, std::ios::out | std::ios::app);
    if (log_file.is_open()) {
        log("=== Debug logging started ===");
    }
}

DebugLogger::~DebugLogger() {
    if (log_file.is_open()) {
        log("=== Debug logging ended ===");
        log_file.close();
    }
}

std::string DebugLogger::getCurrentTimestamp() const {
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(
        now.time_since_epoch()) % 1000;
    
    std::stringstream ss;
    ss << std::put_time(std::localtime(&time_t), "%Y-%m-%d %H:%M:%S");
    ss << '.' << std::setfill('0') << std::setw(3) << ms.count();
    return ss.str();
}

std::string DebugLogger::getThreadId() const {
    std::stringstream ss;
    ss << std::this_thread::get_id();
    return ss.str();
}

void DebugLogger::setEnabled(bool enable) {
    enabled = enable;
}

void DebugLogger::log(const std::string& message) {
    if (!enabled || !log_file.is_open()) return;
    
    std::lock_guard<std::mutex> lock(log_mutex);
    log_file << "[" << getCurrentTimestamp() << "] [" << getThreadId() << "] " 
             << message << std::endl;
    log_file.flush();
}

void DebugLogger::logInfo(const std::string& message) {
    log("[INFO] " + message);
}

void DebugLogger::logWarning(const std::string& message) {
    log("[WARN] " + message);
}

void DebugLogger::logError(const std::string& message) {
    log("[ERROR] " + message);
}

void DebugLogger::logDebug(const std::string& message) {
    log("[DEBUG] " + message);
}

DebugLogger& DebugLogger::getInstance() {
    static DebugLogger instance("debug.log");
    return instance;
}

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

XMLSpreadsheetWriter::XMLSpreadsheetWriter(const std::string& filename) : file(filename) {
    DEBUG_LOG("XMLSpreadsheetWriter: Opening file " + filename);
}

XMLSpreadsheetWriter::~XMLSpreadsheetWriter() {
    if (file.is_open()) {
        DEBUG_LOG("XMLSpreadsheetWriter: Closing file");
        file.close();
    }
}

void XMLSpreadsheetWriter::addWorksheet(const std::string& name) {
    std::string clean_name = cleanSheetName(name);
    worksheets[clean_name] = std::vector<std::vector<std::string>>();
    DEBUG_LOG("XMLSpreadsheetWriter: Added worksheet '" + clean_name + "'");
}

void XMLSpreadsheetWriter::addRow(const std::string& worksheet_name, const std::vector<std::string>& row) {
    std::string clean_name = cleanSheetName(worksheet_name);
    if (worksheets.find(clean_name) != worksheets.end()) {
        worksheets[clean_name].push_back(row);
        DEBUG_LOG("XMLSpreadsheetWriter: Added row to worksheet '" + clean_name + "' (now has " + 
                 std::to_string(worksheets[clean_name].size()) + " rows)");
    } else {
        WARN_LOG("XMLSpreadsheetWriter: Attempted to add row to non-existent worksheet '" + clean_name + "'");
    }
}

bool XMLSpreadsheetWriter::writeFile() {
    DEBUG_LOG("XMLSpreadsheetWriter: Starting to write XML file");
    
    if (!file.is_open()) {
        ERROR_LOG("XMLSpreadsheetWriter: File is not open for writing");
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
    DEBUG_LOG("XMLSpreadsheetWriter: Writing " + std::to_string(worksheets.size()) + " worksheets");
    for (const auto& [sheet_name, rows] : worksheets) {
        DEBUG_LOG("XMLSpreadsheetWriter: Writing worksheet '" + sheet_name + "' with " + std::to_string(rows.size()) + " rows");
        
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
    
    DEBUG_LOG("XMLSpreadsheetWriter: Successfully wrote XML file");
    return true;
}

bool XMLSpreadsheetWriter::isOpen() const {
    return file.is_open();
}

// ProgressTracker implementation
void ProgressTracker::setTotal(int t) {
    total = t;
    start_time = std::chrono::steady_clock::now();
    INFO_LOG("ProgressTracker: Set total files to process: " + std::to_string(t));
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
        INFO_LOG("ProgressTracker: Processing complete - " + std::to_string(proc) + "/" + std::to_string(tot) + " files processed");
    }
}

// RegexAnalyzer implementation
std::vector<ExpressionPattern> RegexAnalyzer::loadExpressions(const std::string& filename) {
    INFO_LOG("RegexAnalyzer: Loading expressions from " + filename);
    
    std::vector<ExpressionPattern> patterns;
    std::ifstream file(filename);
    
    if (!file.is_open()) {
        ERROR_LOG("RegexAnalyzer: Could not open expressions.properties file: " + filename);
        throw std::runtime_error("Could not open expressions.properties file");
    }
    
    std::string line;
    bool in_expressions_section = false;
    int line_num = 0;
    
    while (std::getline(file, line)) {
        line_num++;
        
        // Trim whitespace
        line.erase(0, line.find_first_not_of(" \t\r\n"));
        line.erase(line.find_last_not_of(" \t\r\n") + 1);
        
        // Skip empty lines and comments
        if (line.empty() || line[0] == '#') continue;
        
        // Check for [expressions] section
        if (line == "[expressions]") {
            in_expressions_section = true;
            DEBUG_LOG("RegexAnalyzer: Found [expressions] section at line " + std::to_string(line_num));
            continue;
        }
        
        // Check for other sections
        if (line[0] == '[' && line.back() == ']') {
            in_expressions_section = false;
            DEBUG_LOG("RegexAnalyzer: Left expressions section, found new section: " + line);
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
                    DEBUG_LOG("RegexAnalyzer: Processing expression '" + expr_name + "' = '" + value + "'");
                    
                    try {
                        // Check for inline flags like (?i) at the beginning
                        std::regex_constants::syntax_option_type flags = std::regex_constants::ECMAScript;
                        std::string pattern_str = value;
                        
                        // Handle (?i) case-insensitive flag
                        if (pattern_str.substr(0, 4) == "(?i)") {
                            flags |= std::regex_constants::icase;
                            pattern_str = pattern_str.substr(4); // Remove (?i) from pattern
                            DEBUG_LOG("RegexAnalyzer: Applied case-insensitive flag to " + expr_name);
                        }
                        // Handle (?-i) case-sensitive flag (explicit)
                        else if (pattern_str.substr(0, 5) == "(?-i)") {
                            // Default is case-sensitive, so just remove the flag
                            pattern_str = pattern_str.substr(5); // Remove (?-i) from pattern
                            DEBUG_LOG("RegexAnalyzer: Applied case-sensitive flag to " + expr_name);
                        }
                        // Default case-insensitive behavior (as originally implemented)
                        else {
                            flags |= std::regex_constants::icase;
                            DEBUG_LOG("RegexAnalyzer: Applied default case-insensitive flag to " + expr_name);
                        }
                        
                        std::regex pattern(pattern_str, flags);
                        patterns.push_back({expr_name, std::move(pattern)});
                        std::cout << "Loaded expression: " << expr_name << " = " << value << std::endl;
                        INFO_LOG("RegexAnalyzer: Successfully loaded expression '" + expr_name + "'");
                    } catch (const std::regex_error& e) {
                        std::cerr << "Invalid regex for " << expr_name << ": " << value 
                                 << " Error: " << e.what() << std::endl;
                        ERROR_LOG("RegexAnalyzer: Invalid regex for '" + expr_name + "': " + value + " - " + e.what());
                    }
                }
            }
        }
    }
    
    INFO_LOG("RegexAnalyzer: Loaded " + std::to_string(patterns.size()) + " valid expressions");
    return patterns;
}

bool RegexAnalyzer::isTextFile(const std::string& filepath) {
    try {
        DEBUG_LOG("RegexAnalyzer: Checking if file is text: " + filepath);
        
        std::ifstream file(filepath, std::ios::binary);
        if (!file.is_open()) {
            std::cerr << "Warning: Cannot open file for text check: " << filepath << std::endl;
            WARN_LOG("RegexAnalyzer: Cannot open file for text check: " + filepath);
            return false;
        }
        
        // Read first chunk of file to analyze
        const size_t sample_size = 8192; // 8KB sample
        char buffer[8192];
        
        // Initialize buffer to prevent reading garbage
        memset(buffer, 0, sizeof(buffer));
        
        file.read(buffer, sample_size);
        std::streamsize bytes_read = file.gcount();
        
        if (bytes_read <= 0) {
            DEBUG_LOG("RegexAnalyzer: File is empty, considering as text: " + filepath);
            return true; // Empty file is technically text
        }
        
        // Ensure we don't read beyond what was actually read
        if (bytes_read > static_cast<std::streamsize>(sample_size)) {
            bytes_read = static_cast<std::streamsize>(sample_size);
        }
        
        // Check for null bytes (common in binary files)
        int null_count = 0;
        int printable_count = 0;
        int control_count = 0;
        
        // Only iterate over the actual bytes read with bounds check
        for (std::streamsize i = 0; i < bytes_read && i < static_cast<std::streamsize>(sample_size); ++i) {
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
        if (bytes_read > 0 && null_count > bytes_read * 0.05) {
            DEBUG_LOG("RegexAnalyzer: File has too many null bytes (" + std::to_string(null_count) + "/" + 
                     std::to_string(bytes_read) + "), considering binary: " + filepath);
            return false;
        }
        
        // Heuristic: if less than 70% printable characters, likely binary
        if (bytes_read > 0) {
            double printable_ratio = static_cast<double>(printable_count) / bytes_read;
            if (printable_ratio < 0.70) {
                DEBUG_LOG("RegexAnalyzer: File has low printable ratio (" + std::to_string(printable_ratio) + 
                         "), considering binary: " + filepath);
                return false;
            }
        }
        
        // Additional check for UTF-8 BOM - ensure we have enough bytes
        if (bytes_read >= 3 && 
            static_cast<unsigned char>(buffer[0]) == 0xEF &&
            static_cast<unsigned char>(buffer[1]) == 0xBB &&
            static_cast<unsigned char>(buffer[2]) == 0xBF) {
            DEBUG_LOG("RegexAnalyzer: File has UTF-8 BOM, considering text: " + filepath);
            return true; // UTF-8 BOM indicates text file
        }
        
        DEBUG_LOG("RegexAnalyzer: File passed heuristics, considering text: " + filepath);
        return true; // Passed all heuristics
        
    } catch (const std::exception& e) {
        std::cerr << "Error checking if file is text: " << filepath << " - " << e.what() << std::endl;
        ERROR_LOG("RegexAnalyzer: Error checking if file is text: " + filepath + " - " + e.what());
        return false;
    } catch (...) {
        std::cerr << "Unknown error checking if file is text: " << filepath << std::endl;
        ERROR_LOG("RegexAnalyzer: Unknown error checking if file is text: " + filepath);
        return false;
    }
}

void RegexAnalyzer::processFile(const std::string& filepath) {
    DEBUG_LOG("RegexAnalyzer: Starting to process file: " + filepath);
    
    try {
        std::ifstream file(filepath);
        if (!file.is_open()) {
            std::cerr << "Warning: Could not open file: " << filepath << std::endl;
            WARN_LOG("RegexAnalyzer: Could not open file: " + filepath);
            progress.increment();
            return;
        }
        
        std::vector<Finding> local_findings;
        local_findings.reserve(100);
        
        const size_t BUFFER_SIZE = 64 * 1024; // 64KB buffer
        const size_t WINDOW_SIZE = 32 * 1024; // 32KB sliding window for regex processing
        const size_t OVERLAP_SIZE = 16 * 1024; // 16KB overlap to catch patterns across boundaries
        
        char buffer[BUFFER_SIZE];
        std::string window;
        window.reserve(WINDOW_SIZE + OVERLAP_SIZE);
        int line_number = 1;
        size_t file_position = 0;
        
        DEBUG_LOG("RegexAnalyzer: Processing file with " + std::to_string(expressions.size()) + " expressions: " + filepath);
        
        while (file.read(buffer, BUFFER_SIZE) || file.gcount() > 0) {
            std::streamsize bytes_read = file.gcount();
            
            // Add new data to window
            for (std::streamsize i = 0; i < bytes_read; ++i) {
                window += buffer[i];
                
                // When window gets full, process it and slide
                if (window.size() >= WINDOW_SIZE + OVERLAP_SIZE) {
                    // Process the first WINDOW_SIZE bytes
                    std::string segment = window.substr(0, WINDOW_SIZE);
                    
                    // Count line numbers in this segment
                    int segment_line_start = line_number;
                    for (size_t j = 0; j < segment.size(); ++j) {
                        if (segment[j] == '\n') {
                            line_number++;
                        }
                    }
                    
                    // Process this segment with all expressions
                    for (size_t expr_idx = 0; expr_idx < expressions.size(); ++expr_idx) {
                        try {
                            const auto& expr = expressions[expr_idx];
                            
                            if (expr.name.empty()) {
                                continue;
                            }
                            
                            std::sregex_iterator regex_start(segment.begin(), segment.end(), expr.pattern);
                            std::sregex_iterator regex_end;
                            
                            int matches_in_segment = 0;
                            for (std::sregex_iterator it = regex_start; it != regex_end; ++it) {
                                std::smatch match = *it;
                                matches_in_segment++;
                                
                                // Calculate approximate line number for this match
                                std::string before_match = segment.substr(0, match.position());
                                int match_line = segment_line_start;
                                for (char c : before_match) {
                                    if (c == '\n') match_line++;
                                }
                                
                                // Extract the line containing the match
                                size_t line_start = match.position();
                                while (line_start > 0 && segment[line_start - 1] != '\n') {
                                    line_start--;
                                }
                                
                                size_t line_end = match.position() + match.length();
                                while (line_end < segment.size() && segment[line_end] != '\n') {
                                    line_end++;
                                }
                                
                                std::string match_line_content = segment.substr(line_start, line_end - line_start);
                                
                                Finding finding;
                                finding.expression_name = expr.name;
                                finding.filename = filepath;
                                finding.line_number = match_line;
                                finding.actual_match = match.str();
                                finding.statement = match_line_content;
                                
                                local_findings.push_back(std::move(finding));
                            }
                            
                            if (matches_in_segment > 0) {
                                DEBUG_LOG("RegexAnalyzer: Found " + std::to_string(matches_in_segment) + 
                                         " matches for expression '" + expr.name + "' in segment of " + filepath);
                            }
                            
                        } catch (const std::regex_error& e) {
                            ERROR_LOG("RegexAnalyzer: Regex error processing " + filepath + ": " + e.what());
                            continue;
                        } catch (const std::exception& e) {
                            ERROR_LOG("RegexAnalyzer: Exception processing " + filepath + ": " + e.what());
                            continue;
                        }
                    }
                    
                    // Slide the window - keep the overlap part
                    window = window.substr(WINDOW_SIZE - OVERLAP_SIZE);
                    file_position += WINDOW_SIZE - OVERLAP_SIZE;
                }
            }
        }
        
        // Process any remaining data in the window
        if (!window.empty()) {
            DEBUG_LOG("RegexAnalyzer: Processing final window segment of " + std::to_string(window.size()) + 
                     " bytes for " + filepath);
            
            // Count remaining line numbers
            for (char c : window) {
                if (c == '\n') {
                    line_number++;
                }
            }
            
            // Process remaining content with all expressions
            for (size_t expr_idx = 0; expr_idx < expressions.size(); ++expr_idx) {
                try {
                    const auto& expr = expressions[expr_idx];
                    
                    if (expr.name.empty()) {
                        continue;
                    }
                    
                    std::sregex_iterator regex_start(window.begin(), window.end(), expr.pattern);
                    std::sregex_iterator regex_end;
                    
                    int matches_in_final = 0;
                    for (std::sregex_iterator it = regex_start; it != regex_end; ++it) {
                        std::smatch match = *it;
                        matches_in_final++;
                        
                        // Calculate line number for this match
                        std::string before_match = window.substr(0, match.position());
                        int match_line = line_number - std::count(window.begin(), window.end(), '\n');
                        for (char c : before_match) {
                            if (c == '\n') match_line++;
                        }
                        
                        // Extract the line containing the match
                        size_t line_start = match.position();
                        while (line_start > 0 && window[line_start - 1] != '\n') {
                            line_start--;
                        }
                        
                        size_t line_end = match.position() + match.length();
                        while (line_end < window.size() && window[line_end] != '\n') {
                            line_end++;
                        }
                        
                        std::string match_line_content = window.substr(line_start, line_end - line_start);
                        
                        Finding finding;
                        finding.expression_name = expr.name;
                        finding.filename = filepath;
                        finding.line_number = match_line;
                        finding.actual_match = match.str();
                        finding.statement = match_line_content;
                        
                        local_findings.push_back(std::move(finding));
                    }
                    
                    if (matches_in_final > 0) {
                        DEBUG_LOG("RegexAnalyzer: Found " + std::to_string(matches_in_final) + 
                                 " matches for expression '" + expr.name + "' in final segment of " + filepath);
                    }
                    
                } catch (const std::regex_error& e) {
                    ERROR_LOG("RegexAnalyzer: Regex error in final segment of " + filepath + ": " + e.what());
                    continue;
                } catch (const std::exception& e) {
                    ERROR_LOG("RegexAnalyzer: Exception in final segment of " + filepath + ": " + e.what());
                    continue;
                }
            }
        }
        
        // Add findings
        if (!local_findings.empty()) {
            std::lock_guard<std::mutex> lock(findings_mutex);
            all_findings.insert(all_findings.end(), 
                               std::make_move_iterator(local_findings.begin()),
                               std::make_move_iterator(local_findings.end()));
            INFO_LOG("RegexAnalyzer: Added " + std::to_string(local_findings.size()) + 
                    " findings from " + filepath + " (total now: " + std::to_string(all_findings.size()) + ")");
        } else {
            DEBUG_LOG("RegexAnalyzer: No findings in " + filepath);
        }
        
    } catch (const std::exception& e) {
        std::cerr << "Fatal error processing file " << filepath << ": " << e.what() << std::endl;
        ERROR_LOG("RegexAnalyzer: Fatal error processing file " + filepath + ": " + e.what());
    } catch (...) {
        std::cerr << "Unknown fatal error processing file " << filepath << std::endl;
        ERROR_LOG("RegexAnalyzer: Unknown fatal error processing file " + filepath);
    }
    
    DEBUG_LOG("RegexAnalyzer: Completed processing file: " + filepath);
    progress.increment();
}

std::vector<std::string> RegexAnalyzer::findTextFiles(const std::string& directory) {
    INFO_LOG("RegexAnalyzer: Starting to find text files in directory: " + directory);
    
    std::vector<std::string> text_files;
    
    try {
        if (!std::filesystem::exists(directory)) {
            std::cerr << "Error: Directory does not exist: " << directory << std::endl;
            ERROR_LOG("RegexAnalyzer: Directory does not exist: " + directory);
            return text_files;
        }
        
        if (!std::filesystem::is_directory(directory)) {
            std::cerr << "Error: Path is not a directory: " << directory << std::endl;
            ERROR_LOG("RegexAnalyzer: Path is not a directory: " + directory);
            return text_files;
        }
        
        int total_files = 0;
        int text_file_count = 0;
        
        for (const auto& entry : std::filesystem::recursive_directory_iterator(directory)) {
            try {
                if (entry.is_regular_file()) {
                    total_files++;
                    if (isTextFile(entry.path().string())) {
                        text_files.push_back(entry.path().string());
                        text_file_count++;
                        
                        if (text_file_count % 100 == 0) {
                            DEBUG_LOG("RegexAnalyzer: Found " + std::to_string(text_file_count) + " text files so far...");
                        }
                    }
                }
            } catch (const std::filesystem::filesystem_error& e) {
                std::cerr << "Error accessing file: " << entry.path() 
                         << " - " << e.what() << std::endl;
                WARN_LOG("RegexAnalyzer: Error accessing file: " + entry.path().string() + " - " + e.what());
                continue; // Skip this file and continue
            }
        }
        
        INFO_LOG("RegexAnalyzer: Scanned " + std::to_string(total_files) + " total files, found " + 
                std::to_string(text_file_count) + " text files");
        
    } catch (const std::filesystem::filesystem_error& e) {
        std::cerr << "Error accessing directory: " << e.what() << std::endl;
        ERROR_LOG("RegexAnalyzer: Error accessing directory: " + std::string(e.what()));
    }
    
    return text_files;
}

void RegexAnalyzer::workerThread() {
    DEBUG_LOG("RegexAnalyzer: Worker thread started");
    
    int files_processed = 0;
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
        
        // Debug output to track which file is being processed
        static std::mutex debug_mutex;
        {
            std::lock_guard<std::mutex> lock(debug_mutex);
            std::cout << "Processing: " << filepath << std::endl;
        }
        
        processFile(filepath);
        files_processed++;
    }
    
    DEBUG_LOG("RegexAnalyzer: Worker thread completed, processed " + std::to_string(files_processed) + " files");
}

void RegexAnalyzer::analyze(const std::string& directory, const std::string& expressions_file, 
            const std::string& output_file, int num_threads) {
    
    INFO_LOG("RegexAnalyzer: Starting analysis - directory: " + directory + 
            ", expressions: " + expressions_file + ", output: " + output_file + 
            ", threads: " + std::to_string(num_threads));
    
    std::cout << "Loading expressions from: " << expressions_file << std::endl;
    expressions = loadExpressions(expressions_file);
    
    if (expressions.empty()) {
        ERROR_LOG("RegexAnalyzer: No valid expressions found in properties file");
        throw std::runtime_error("No valid expressions found in properties file");
    }
    
    std::cout << "Loaded " << expressions.size() << " expressions" << std::endl;
    std::cout << "Scanning directory: " << directory << std::endl;
    
    file_queue = findTextFiles(directory);
    std::cout << "Found " << file_queue.size() << " text files" << std::endl;
    
    if (file_queue.empty()) {
        std::cout << "No text files found to process" << std::endl;
        WARN_LOG("RegexAnalyzer: No text files found to process");
        return;
    }
    
    progress.setTotal(file_queue.size());
    std::cout << "Starting analysis with " << num_threads << " threads..." << std::endl;
    
    auto start_time = std::chrono::steady_clock::now();
    
    // Launch worker threads
    std::vector<std::thread> threads;
    for (int i = 0; i < num_threads; ++i) {
        threads.emplace_back(&RegexAnalyzer::workerThread, this);
        DEBUG_LOG("RegexAnalyzer: Launched worker thread " + std::to_string(i + 1));
    }
    
    // Wait for all threads to complete
    for (auto& thread : threads) {
        thread.join();
    }
    
    auto end_time = std::chrono::steady_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::seconds>(end_time - start_time);
    
    std::cout << std::endl << "Analysis complete. Found " << all_findings.size() << " matches" << std::endl;
    std::cout << "Writing results to: " << output_file << std::endl;
    
    INFO_LOG("RegexAnalyzer: Analysis completed in " + std::to_string(duration.count()) + 
            " seconds, found " + std::to_string(all_findings.size()) + " total matches");
    
    writeResults(output_file);
    
    INFO_LOG("RegexAnalyzer: Analysis fully completed");
}

void RegexAnalyzer::writeResults(const std::string& output_filename) {
    INFO_LOG("RegexAnalyzer: Starting to write results to " + output_filename);
    
#if USE_XLSX
    writeXLSXResults(output_filename);
#else
    writeXMLSpreadsheetResults(output_filename);
#endif
}

#if USE_XLSX
void RegexAnalyzer::writeXLSXResults(const std::string& output_filename) {
    INFO_LOG("RegexAnalyzer: Writing XLSX results to " + output_filename);
    
    // Create workbook
    lxw_workbook* workbook = workbook_new(output_filename.c_str());
    if (!workbook) {
        ERROR_LOG("RegexAnalyzer: Failed to create Excel workbook: " + output_filename);
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
    
    DEBUG_LOG("RegexAnalyzer: Grouped findings into " + std::to_string(grouped_findings.size()) + " expression categories");
    
    // Create worksheet for each expression
    for (const auto& [expr_name, findings] : grouped_findings) {
        DEBUG_LOG("RegexAnalyzer: Creating XLSX worksheet for expression '" + expr_name + 
                 "' with " + std::to_string(findings.size()) + " findings");
        
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
            ERROR_LOG("RegexAnalyzer: Failed to create worksheet: " + sheet_name);
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
        DEBUG_LOG("RegexAnalyzer: Creating XLSX summary worksheet with " + std::to_string(all_findings.size()) + " findings");
        
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
        ERROR_LOG("RegexAnalyzer: Failed to save Excel workbook: " + std::string(lxw_strerror(error)));
        throw std::runtime_error("Failed to save Excel workbook: " + std::string(lxw_strerror(error)));
    }
    
    std::cout << "Successfully created Excel file: " << output_filename << std::endl;
    INFO_LOG("RegexAnalyzer: Successfully created XLSX file: " + output_filename);
}
#endif

void RegexAnalyzer::writeXMLSpreadsheetResults(const std::string& output_filename) {
    INFO_LOG("RegexAnalyzer: Writing XML Spreadsheet results to " + output_filename);
    
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
    
    DEBUG_LOG("RegexAnalyzer: Final XML filename: " + xml_filename);
    
    XMLSpreadsheetWriter writer(xml_filename);
    if (!writer.isOpen()) {
        ERROR_LOG("RegexAnalyzer: Failed to create XML spreadsheet: " + xml_filename);
        throw std::runtime_error("Failed to create XML spreadsheet: " + xml_filename);
    }
    
    // Group findings by expression
    std::map<std::string, std::vector<Finding>> grouped_findings;
    
    for (const auto& finding : all_findings) {
        grouped_findings[finding.expression_name].push_back(finding);
    }
    
    DEBUG_LOG("RegexAnalyzer: Grouped findings into " + std::to_string(grouped_findings.size()) + " expression categories");
    
    // Create worksheet for each expression
    for (const auto& [expr_name, findings] : grouped_findings) {
        DEBUG_LOG("RegexAnalyzer: Creating XML worksheet for expression '" + expr_name + 
                 "' with " + std::to_string(findings.size()) + " findings");
        
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
        DEBUG_LOG("RegexAnalyzer: Creating XML summary worksheet with " + std::to_string(all_findings.size()) + " findings");
        
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
        ERROR_LOG("RegexAnalyzer: Failed to write XML spreadsheet file");
        throw std::runtime_error("Failed to write XML spreadsheet file");
    }
    
    std::cout << "Successfully created XML Spreadsheet file: " << xml_filename << std::endl;
    std::cout << "This file can be opened in Excel, LibreOffice Calc, or Google Sheets" << std::endl;
    
    INFO_LOG("RegexAnalyzer: Successfully created XML Spreadsheet file: " + xml_filename);
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
    INFO_LOG("Program started with " + std::to_string(argc) + " arguments");
    
    if (argc < 4 || argc > 5) {
        printUsage(argv[0]);
        ERROR_LOG("Invalid number of arguments provided");
        return 1;
    }
    
    std::string directory = argv[1];
    std::string expressions_file = argv[2];
    std::string output_file = argv[3];
    int num_threads = (argc == 5) ? std::stoi(argv[4]) : 4;
    
    INFO_LOG("Arguments - Directory: " + directory + ", Expressions: " + expressions_file + 
            ", Output: " + output_file + ", Threads: " + std::to_string(num_threads));
    
#if USE_XLSX
    std::cout << "Using XLSX output format" << std::endl;
    INFO_LOG("Using XLSX output format");
#else
    std::cout << "Using XML Spreadsheet 2003 output format (XLSX library not available)" << std::endl;
    INFO_LOG("Using XML Spreadsheet 2003 output format (XLSX library not available)");
#endif
    
    try {
        RegexAnalyzer analyzer;
        analyzer.analyze(directory, expressions_file, output_file, num_threads);
        
        std::cout << "Analysis completed successfully!" << std::endl;
        INFO_LOG("Analysis completed successfully!");
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        ERROR_LOG("Fatal error: " + std::string(e.what()));
        return 1;
    }
}
