#include "whistle.h"

// XMLSpreadsheetWriter implementation (unchanged)
std::string XMLSpreadsheetWriter::escapeXML(const std::string& text) {
    if (text.empty()) {
        return text;
    }
    
    std::string escaped;
    escaped.reserve(text.length() * 2);
    
    for (size_t i = 0; i < text.length(); ++i) {
        char c = text[i];
        switch (c) {
            case '&': escaped += "&amp;"; break;
            case '<': escaped += "&lt;"; break;
            case '>': escaped += "&gt;"; break;
            case '"': escaped += "&quot;"; break;
            case '\'': escaped += "&apos;"; break;
            default: escaped += c; break;
        }
    }
    return escaped;
}

std::string XMLSpreadsheetWriter::cleanSheetName(const std::string& name) {
    if (name.empty()) return "Sheet1";
    
    std::string clean = name;
    for (size_t i = 0; i < clean.length(); ++i) {
        char& c = clean[i];
        if (c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']' || c == ':') {
            c = '_';
        }
    }
    
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
    if (!file.is_open()) return false;
    
    // Write XML header
    file << "<?xml version=\"1.0\"?>\n";
    file << "<?mso-application progid=\"Excel.Sheet\"?>\n";
    file << "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n";
    file << " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n";
    file << " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n";
    file << " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n";
    file << " xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n";
    
    // Write document properties
    file << " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n";
    file << "  <Created>" << std::chrono::system_clock::now().time_since_epoch().count() << "</Created>\n";
    file << "  <Application>Regex Analyzer</Application>\n";
    file << " </DocumentProperties>\n";
    
    // Write styles (header and cell styles)
    file << " <Styles>\n";
    file << "  <Style ss:ID=\"Header\">\n";
    file << "   <Font ss:Bold=\"1\"/>\n";
    file << "   <Interior ss:Color=\"#C0C0C0\" ss:Pattern=\"Solid\"/>\n";
    file << "   <Borders>\n";
    file << "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "   </Borders>\n";
    file << "  </Style>\n";
    file << "  <Style ss:ID=\"Cell\">\n";
    file << "   <Borders>\n";
    file << "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\n";
    file << "   </Borders>\n";
    file << "   <Alignment ss:Vertical=\"Top\" ss:WrapText=\"1\"/>\n";
    file << "  </Style>\n";
    file << " </Styles>\n";
    
    // Write worksheets
    for (const auto& [sheet_name, rows] : worksheets) {
        file << " <Worksheet ss:Name=\"" << escapeXML(sheet_name) << "\">\n";
        file << "  <Table>\n";
        
        // Set column widths
        file << "   <Column ss:Width=\"120\"/>\n"; // Finding
        file << "   <Column ss:Width=\"240\"/>\n"; // File
        file << "   <Column ss:Width=\"60\"/>\n";  // Line
        file << "   <Column ss:Width=\"120\"/>\n"; // Comments
        file << "   <Column ss:Width=\"90\"/>\n";  // Ease
        file << "   <Column ss:Width=\"90\"/>\n";  // Significance
        file << "   <Column ss:Width=\"90\"/>\n";  // Risk
        file << "   <Column ss:Width=\"360\"/>\n"; // Statement
        
        for (size_t i = 0; i < rows.size(); ++i) {
            const auto& row = rows[i];
            file << "   <Row>\n";
            
            for (size_t j = 0; j < row.size(); ++j) {
                std::string style_id = (i == 0) ? "Header" : "Cell";
                std::string cell_data = escapeXML(row[j]);
                
                bool is_number = false;
                if (j == 2 && i > 0) { // Line number column
                    try {
                        std::stoi(row[j]);
                        is_number = true;
                    } catch (...) {
                        is_number = false;
                    }
                }
                
                file << "    <Cell ss:StyleID=\"" << style_id << "\">\n";
                if (is_number) {
                    file << "     <Data ss:Type=\"Number\">" << cell_data << "</Data>\n";
                } else {
                    file << "     <Data ss:Type=\"String\">" << cell_data << "</Data>\n";
                }
                file << "    </Cell>\n";
            }
            file << "   </Row>\n";
        }
        
        file << "  </Table>\n";
        
        // Add worksheet options (freeze header row)
        if (!rows.empty()) {
            file << "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n";
            file << "   <FreezePanes/>\n";
            file << "   <FrozenNoSplit/>\n";
            file << "   <SplitHorizontal>1</SplitHorizontal>\n";
            file << "   <TopRowBottomPane>1</TopRowBottomPane>\n";
            file << "   <ActivePane>2</ActivePane>\n";
            file << "  </WorksheetOptions>\n";
        }
        
        file << " </Worksheet>\n";
    }
    
    file << "</Workbook>\n";
    return true;
}

bool XMLSpreadsheetWriter::isOpen() const {
    return file.is_open();
}

// ProgressTracker implementation (unchanged)
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

// AsyncRegexAnalyzer implementation
std::vector<ExpressionPattern> AsyncRegexAnalyzer::loadExpressions(const std::string& filename) {
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
                        std::regex_constants::syntax_option_type flags = std::regex_constants::ECMAScript;
                        std::string pattern_str = value;
                        
                        // Handle (?i) case-insensitive flag
                        if (pattern_str.substr(0, 4) == "(?i)") {
                            flags |= std::regex_constants::icase;
                            pattern_str = pattern_str.substr(4);
                        } else if (pattern_str.substr(0, 5) == "(?-i)") {
                            pattern_str = pattern_str.substr(5);
                        } else {
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

bool AsyncRegexAnalyzer::isTextFile(const std::string& filepath) {
    try {
        std::ifstream file(filepath, std::ios::binary);
        if (!file.is_open()) {
            std::cerr << "Warning: Cannot open file for text check: " << filepath << std::endl;
            return false;
        }
        
        const size_t sample_size = 8192;
        char buffer[8192];
        memset(buffer, 0, sizeof(buffer));
        
        file.read(buffer, sample_size);
        std::streamsize bytes_read = file.gcount();
        
        if (bytes_read <= 0) return true;
        
        if (bytes_read > static_cast<std::streamsize>(sample_size)) {
            bytes_read = static_cast<std::streamsize>(sample_size);
        }
        
        int null_count = 0;
        int printable_count = 0;
        
        for (std::streamsize i = 0; i < bytes_read && i < static_cast<std::streamsize>(sample_size); ++i) {
            unsigned char byte = static_cast<unsigned char>(buffer[i]);
            
            if (byte == 0) {
                null_count++;
            } else if (std::isprint(byte) || byte == '\t' || byte == '\n' || byte == '\r') {
                printable_count++;
            }
        }
        
        if (bytes_read > 0 && null_count > bytes_read * 0.05) {
            return false;
        }
        
        if (bytes_read > 0) {
            double printable_ratio = static_cast<double>(printable_count) / bytes_read;
            if (printable_ratio < 0.70) {
                return false;
            }
        }
        
        // Check for UTF-8 BOM
        if (bytes_read >= 3 && 
            static_cast<unsigned char>(buffer[0]) == 0xEF &&
            static_cast<unsigned char>(buffer[1]) == 0xBB &&
            static_cast<unsigned char>(buffer[2]) == 0xBF) {
            return true;
        }
        
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "Error checking if file is text: " << filepath << " - " << e.what() << std::endl;
        return false;
    } catch (...) {
        std::cerr << "Unknown error checking if file is text: " << filepath << std::endl;
        return false;
    }
}

std::vector<Finding> AsyncRegexAnalyzer::processFileWithExpression(const std::string& filepath, 
                                                                   const ExpressionPattern& expression) {
    std::vector<Finding> findings;
    
    try {
        std::ifstream file(filepath);
        if (!file.is_open()) {
            return findings; // Return empty findings
        }
        
        const size_t BUFFER_SIZE = 64 * 1024;
        const size_t WINDOW_SIZE = 32 * 1024;
        const size_t OVERLAP_SIZE = 16 * 1024;
        
        char buffer[BUFFER_SIZE];
        std::string window;
        window.reserve(WINDOW_SIZE + OVERLAP_SIZE);
        int line_number = 1;
        
        while (file.read(buffer, BUFFER_SIZE) || file.gcount() > 0) {
            std::streamsize bytes_read = file.gcount();
            
            for (std::streamsize i = 0; i < bytes_read; ++i) {
                window += buffer[i];
                
                if (window.size() >= WINDOW_SIZE + OVERLAP_SIZE) {
                    std::string segment = window.substr(0, WINDOW_SIZE);
                    
                    int segment_line_start = line_number;
                    for (size_t j = 0; j < segment.size(); ++j) {
                        if (segment[j] == '\n') {
                            line_number++;
                        }
                    }
                    
                    // Process this segment with the expression
                    try {
                        std::sregex_iterator regex_start(segment.begin(), segment.end(), expression.pattern);
                        std::sregex_iterator regex_end;
                        
                        for (std::sregex_iterator it = regex_start; it != regex_end; ++it) {
                            std::smatch match = *it;
                            
                            // Calculate line number
                            std::string before_match = segment.substr(0, match.position());
                            int match_line = segment_line_start;
                            for (char c : before_match) {
                                if (c == '\n') match_line++;
                            }
                            
                            // Extract line containing match
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
                            finding.expression_name = expression.name;
                            finding.filename = filepath;
                            finding.line_number = match_line;
                            finding.actual_match = match.str();
                            finding.statement = match_line_content;
                            
                            findings.push_back(std::move(finding));
                        }
                    } catch (const std::regex_error& e) {
                        // Continue processing
                    }
                    
                    window = window.substr(WINDOW_SIZE - OVERLAP_SIZE);
                }
            }
        }
        
        // Process remaining data in window
        if (!window.empty()) {
            for (char c : window) {
                if (c == '\n') {
                    line_number++;
                }
            }
            
            try {
                std::sregex_iterator regex_start(window.begin(), window.end(), expression.pattern);
                std::sregex_iterator regex_end;
                
                for (std::sregex_iterator it = regex_start; it != regex_end; ++it) {
                    std::smatch match = *it;
                    
                    std::string before_match = window.substr(0, match.position());
                    int match_line = line_number - std::count(window.begin(), window.end(), '\n');
                    for (char c : before_match) {
                        if (c == '\n') match_line++;
                    }
                    
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
                    finding.expression_name = expression.name;
                    finding.filename = filepath;
                    finding.line_number = match_line;
                    finding.actual_match = match.str();
                    finding.statement = match_line_content;
                    
                    findings.push_back(std::move(finding));
                }
            } catch (const std::regex_error& e) {
                // Continue
            }
        }
        
    } catch (const std::exception& e) {
        std::cerr << "Error processing file " << filepath << " with expression " 
                  << expression.name << ": " << e.what() << std::endl;
    }
    
    return findings;
}

std::future<std::vector<Finding>> AsyncRegexAnalyzer::processFileAsync(const std::string& filepath, 
                                                                       const ExpressionPattern& expression) {
    return std::async(std::launch::async, [this, filepath, expression]() {
        auto findings = processFileWithExpression(filepath, expression);
        progress.increment(); // Increment after each file-expression pair completes
        return findings;
    });
}

std::vector<std::string> AsyncRegexAnalyzer::findTextFiles(const std::string& directory) {
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
                continue;
            }
        }
    } catch (const std::filesystem::filesystem_error& e) {
        std::cerr << "Error accessing directory: " << e.what() << std::endl;
    }
    
    return text_files;
}

void AsyncRegexAnalyzer::collectCompletedFutures() {
    std::lock_guard<std::mutex> lock(futures_mutex);
    
    // Check for completed futures and collect their results
    auto it = file_futures.begin();
    while (it != file_futures.end()) {
        if (it->wait_for(std::chrono::milliseconds(0)) == std::future_status::ready) {
            try {
                auto findings = it->get();
                if (!findings.empty()) {
                    std::lock_guard<std::mutex> findings_lock(findings_mutex);
                    all_findings.insert(all_findings.end(), 
                                       std::make_move_iterator(findings.begin()),
                                       std::make_move_iterator(findings.end()));
                }
            } catch (const std::exception& e) {
                std::cerr << "Error collecting future result: " << e.what() << std::endl;
            }
            it = file_futures.erase(it);
        } else {
            ++it;
        }
    }
}

void AsyncRegexAnalyzer::workerThread() {
    while (!shutdown.load()) {
        WorkItem work_item("", 0);
        
        {
            std::unique_lock<std::mutex> lock(queue_mutex);
            queue_cv.wait(lock, [this] { return !work_queue.empty() || shutdown.load(); });
            
            if (shutdown.load() && work_queue.empty()) {
                break;
            }
            
            if (!work_queue.empty()) {
                work_item = work_queue.front();
                work_queue.pop();
            } else {
                continue;
            }
        }
        
        if (!work_item.filepath.empty() && work_item.expression_index < expressions.size()) {
            // Launch async processing for this file-expression pair
            auto future = processFileAsync(work_item.filepath, expressions[work_item.expression_index]);
            
            std::lock_guard<std::mutex> futures_lock(futures_mutex);
            file_futures.push_back(std::move(future));
        }
        
        // Periodically collect completed futures
        collectCompletedFutures();
    }
}

void AsyncRegexAnalyzer::analyze(const std::string& directory, const std::string& expressions_file, 
                                const std::string& output_file, int num_threads) {
    
    std::cout << "Loading expressions from: " << expressions_file << std::endl;
    expressions = loadExpressions(expressions_file);
    
    if (expressions.empty()) {
        throw std::runtime_error("No valid expressions found in properties file");
    }
    
    std::cout << "Loaded " << expressions.size() << " expressions" << std::endl;
    std::cout << "Scanning directory: " << directory << std::endl;
    
    auto text_files = findTextFiles(directory);
    std::cout << "Found " << text_files.size() << " text files" << std::endl;
    
    if (text_files.empty()) {
        std::cout << "No text files found to process" << std::endl;
        return;
    }
    
    // Create work items for each file-expression combination
    {
        std::lock_guard<std::mutex> lock(queue_mutex);
        for (const auto& filepath : text_files) {
            for (size_t expr_idx = 0; expr_idx < expressions.size(); ++expr_idx) {
                work_queue.emplace(filepath, expr_idx);
            }
        }
    }
    
    int total_work_items = text_files.size() * expressions.size();
    progress.setTotal(total_work_items);
    std::cout << "Created " << total_work_items << " work items (" 
              << text_files.size() << " files × " << expressions.size() << " expressions)" << std::endl;
    std::cout << "Starting analysis with " << num_threads << " threads..." << std::endl;
    
    // Launch worker threads
    std::vector<std::thread> threads;
    for (int i = 0; i < num_threads; ++i) {
        threads.emplace_back(&AsyncRegexAnalyzer::workerThread, this);
    }
    
    // Monitor progress and collect results
    auto start_time = std::chrono::steady_clock::now();
    while (true) {
        std::this_thread::sleep_for(std::chrono::milliseconds(500));
        
        // Collect completed futures
        collectCompletedFutures();
        
        // Check if all work is done
        bool work_queue_empty = false;
        {
            std::lock_guard<std::mutex> lock(queue_mutex);
            work_queue_empty = work_queue.empty();
        }
        
        bool futures_empty = false;
        {
            std::lock_guard<std::mutex> lock(futures_mutex);
            futures_empty = file_futures.empty();
        }
        
        if (work_queue_empty && futures_empty) {
            break;
        }
        
        // Timeout check (prevent infinite wait)
        auto elapsed = std::chrono::duration_cast<std::chrono::minutes>(
            std::chrono::steady_clock::now() - start_time).count();
        if (elapsed > 60) { // 60 minute timeout
            std::cout << "\nWarning: Processing taking longer than expected. Checking for stuck threads..." << std::endl;
        }
    }
    
    // Signal shutdown and wait for threads
    shutdown.store(true);
    queue_cv.notify_all();
    
    for (auto& thread : threads) {
        thread.join();
    }
    
    // Final collection of any remaining futures
    collectCompletedFutures();
    
    std::cout << std::endl << "Analysis complete. Found " << all_findings.size() << " matches" << std::endl;
    std::cout << "Writing results to: " << output_file << std::endl;
    
    writeResults(output_file);
}

void AsyncRegexAnalyzer::writeResults(const std::string& output_filename) {
#if USE_XLSX
    writeXLSXResults(output_filename);
#else
    writeXMLSpreadsheetResults(output_filename);
#endif
}

#if USE_XLSX
void AsyncRegexAnalyzer::writeXLSXResults(const std::string& output_filename) {
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
        std::string sheet_name = expr_name;
        for (char& c : sheet_name) {
            if (c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']' || c == ':') {
                c = '_';
            }
        }
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
            worksheet_write_string(worksheet, row, 0, finding.actual_match.c_str(), cell_format);
            worksheet_write_string(worksheet, row, 1, finding.filename.c_str(), cell_format);
            worksheet_write_number(worksheet, row, 2, finding.line_number, cell_format);
            worksheet_write_string(worksheet, row, 3, "", cell_format); // Comments (blank)
            worksheet_write_string(worksheet, row, 4, "", cell_format); // Ease (blank)
            worksheet_write_string(worksheet, row, 5, "", cell_format); // Significance (blank)
            worksheet_write_string(worksheet, row, 6, "", cell_format); // Risk (blank)
            worksheet_write_string(worksheet, row, 7, finding.statement.c_str(), cell_format);
            row++;
        }
        
        worksheet_freeze_panes(worksheet, 1, 0);
        std::cout << "Created sheet: " << sheet_name << " with " << findings.size() << " findings" << std::endl;
    }
    
    // Create summary worksheet
    if (!all_findings.empty()) {
        lxw_worksheet* summary_worksheet = workbook_add_worksheet(workbook, "Summary");
        if (summary_worksheet) {
            // Set column widths
            worksheet_set_column(summary_worksheet, 0, 0, 20, nullptr);
            worksheet_set_column(summary_worksheet, 1, 1, 40, nullptr);
            worksheet_set_column(summary_worksheet, 2, 2, 10, nullptr);
            worksheet_set_column(summary_worksheet, 3, 3, 20, nullptr);
            worksheet_set_column(summary_worksheet, 4, 4, 15, nullptr);
            worksheet_set_column(summary_worksheet, 5, 5, 15, nullptr);
            worksheet_set_column(summary_worksheet, 6, 6, 15, nullptr);
            worksheet_set_column(summary_worksheet, 7, 7, 60, nullptr);
            
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
                worksheet_write_string(summary_worksheet, row, 0, finding.actual_match.c_str(), cell_format);
                worksheet_write_string(summary_worksheet, row, 1, finding.filename.c_str(), cell_format);
                worksheet_write_number(summary_worksheet, row, 2, finding.line_number, cell_format);
                worksheet_write_string(summary_worksheet, row, 3, "", cell_format);
                worksheet_write_string(summary_worksheet, row, 4, "", cell_format);
                worksheet_write_string(summary_worksheet, row, 5, "", cell_format);
                worksheet_write_string(summary_worksheet, row, 6, "", cell_format);
                worksheet_write_string(summary_worksheet, row, 7, finding.statement.c_str(), cell_format);
                row++;
            }
            
            worksheet_freeze_panes(summary_worksheet, 1, 0);
            std::cout << "Created Summary sheet with " << all_findings.size() << " total findings" << std::endl;
        }
    }
    
    lxw_error error = workbook_close(workbook);
    if (error != LXW_NO_ERROR) {
        throw std::runtime_error("Failed to save Excel workbook: " + std::string(lxw_strerror(error)));
    }
    
    std::cout << "Successfully created Excel file: " << output_filename << std::endl;
}
#endif

void AsyncRegexAnalyzer::writeXMLSpreadsheetResults(const std::string& output_filename) {
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
                finding.actual_match,
                finding.filename,
                std::to_string(finding.line_number),
                "", "", "", "", // Comments, Ease, Significance, Risk (blank)
                finding.statement
            });
        }
        
        std::cout << "Created sheet: " << expr_name << " with " << findings.size() << " findings" << std::endl;
    }
    
    // Create summary worksheet
    if (!all_findings.empty()) {
        writer.addWorksheet("Summary");
        
        writer.addRow("Summary", {"Finding", "File", "Line", "Comments", "Ease", "Significance", "Risk", "Statement"});
        
        for (const auto& finding : all_findings) {
            writer.addRow("Summary", {
                finding.actual_match,
                finding.filename,
                std::to_string(finding.line_number),
                "", "", "", "",
                finding.statement
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
        AsyncRegexAnalyzer analyzer;
        analyzer.analyze(directory, expressions_file, output_file, num_threads);
        
        std::cout << "Analysis completed successfully!" << std::endl;
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }
}
