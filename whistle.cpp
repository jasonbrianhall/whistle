#include <iostream>
#include <fstream>
#include <filesystem>
#include <regex>
#include <vector>
#include <string>
#include <thread>
#include <mutex>
#include <queue>
#include <atomic>
#include <chrono>
#include <map>
#include <iomanip>
#include <sstream>
#include <xlsxwriter.h>

struct Finding {
    std::string expression_name;
    std::string filename;
    int line_number;
    std::string statement;
};

struct ExpressionPattern {
    std::string name;
    std::regex pattern;
};

class ProgressTracker {
private:
    std::atomic<int> processed{0};
    std::atomic<int> total{0};
    std::chrono::steady_clock::time_point start_time;
    mutable std::mutex print_mutex;
    
public:
    void setTotal(int t) {
        total = t;
        start_time = std::chrono::steady_clock::now();
    }
    
    void increment() {
        processed++;
        printProgress();
    }
    
    void printProgress() const {
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
};

class RegexAnalyzer {
private:
    std::vector<ExpressionPattern> expressions;
    std::vector<std::string> file_queue;
    std::mutex queue_mutex;
    std::mutex findings_mutex;
    std::vector<Finding> all_findings;
    ProgressTracker progress;
    
    std::vector<ExpressionPattern> loadExpressions(const std::string& filename) {
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
                            std::regex pattern(value, std::regex_constants::icase);
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
    
    bool isTextFile(const std::string& filepath) {
        std::ifstream file(filepath, std::ios::binary);
        if (!file.is_open()) {
            return false;
        }
        
        // Read first chunk of file to analyze
        const size_t sample_size = 8192; // 8KB sample
        std::vector<char> buffer(sample_size);
        file.read(buffer.data(), sample_size);
        std::streamsize bytes_read = file.gcount();
        
        if (bytes_read == 0) {
            return false; // Empty file
        }
        
        // Check for null bytes (common in binary files)
        int null_count = 0;
        int printable_count = 0;
        int control_count = 0;
        
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
        
        // Additional check for UTF-8 BOM
        if (bytes_read >= 3 && 
            static_cast<unsigned char>(buffer[0]) == 0xEF &&
            static_cast<unsigned char>(buffer[1]) == 0xBB &&
            static_cast<unsigned char>(buffer[2]) == 0xBF) {
            return true; // UTF-8 BOM indicates text file
        }
        
        return true; // Passed all heuristics
    }
    
    std::vector<std::string> findTextFiles(const std::string& directory) {
        std::vector<std::string> text_files;
        
        try {
            for (const auto& entry : std::filesystem::recursive_directory_iterator(directory)) {
                if (entry.is_regular_file()) {
                    // Skip very large files (>100MB) to avoid memory issues
                    if (entry.file_size() > 100 * 1024 * 1024) {
                        continue;
                    }
                    
                    if (isTextFile(entry.path().string())) {
                        text_files.push_back(entry.path().string());
                    }
                }
            }
        } catch (const std::filesystem::filesystem_error& e) {
            std::cerr << "Error accessing directory: " << e.what() << std::endl;
        }
        
        return text_files;
    }
    
    void processFile(const std::string& filepath) {
        std::ifstream file(filepath);
        if (!file.is_open()) {
            return;
        }
        
        std::string line;
        int line_number = 0;
        std::vector<Finding> local_findings;
        
        while (std::getline(file, line)) {
            line_number++;
            
            for (const auto& expr : expressions) {
                std::smatch match;
                if (std::regex_search(line, match, expr.pattern)) {
                    Finding finding;
                    finding.expression_name = expr.name;
                    finding.filename = filepath;
                    finding.line_number = line_number;
                    finding.statement = line;
                    local_findings.push_back(finding);
                }
            }
        }
        
        if (!local_findings.empty()) {
            std::lock_guard<std::mutex> lock(findings_mutex);
            all_findings.insert(all_findings.end(), local_findings.begin(), local_findings.end());
        }
        
        progress.increment();
    }
    
    void workerThread() {
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
    
public:
    void analyze(const std::string& directory, const std::string& expressions_file, 
                const std::string& output_file, int num_threads = 4) {
        
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
    
private:
    void writeResults(const std::string& output_filename) {
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
                worksheet_write_string(worksheet, row, 0, finding.expression_name.c_str(), cell_format);
                worksheet_write_string(worksheet, row, 1, finding.filename.c_str(), cell_format);
                worksheet_write_number(worksheet, row, 2, finding.line_number, cell_format);
                worksheet_write_string(worksheet, row, 3, "", cell_format); // Comments (blank)
                worksheet_write_string(worksheet, row, 4, "", cell_format); // Ease (blank)
                worksheet_write_string(worksheet, row, 5, "", cell_format); // Significance (blank)
                worksheet_write_string(worksheet, row, 6, "", cell_format); // Risk (blank)
                worksheet_write_string(worksheet, row, 7, finding.statement.c_str(), cell_format);
                row++;
            }
            
            // Freeze the header row
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
                    worksheet_write_string(summary_worksheet, row, 0, finding.expression_name.c_str(), cell_format);
                    worksheet_write_string(summary_worksheet, row, 1, finding.filename.c_str(), cell_format);
                    worksheet_write_number(summary_worksheet, row, 2, finding.line_number, cell_format);
                    worksheet_write_string(summary_worksheet, row, 3, "", cell_format); // Comments (blank)
                    worksheet_write_string(summary_worksheet, row, 4, "", cell_format); // Ease (blank)
                    worksheet_write_string(summary_worksheet, row, 5, "", cell_format); // Significance (blank)
                    worksheet_write_string(summary_worksheet, row, 6, "", cell_format); // Risk (blank)
                    worksheet_write_string(summary_worksheet, row, 7, finding.statement.c_str(), cell_format);
                    row++;
                }
                
                // Freeze the header row
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
};

void printUsage(const char* program_name) {
    std::cout << "Usage: " << program_name << " <directory> <expressions_file> <output_file> [num_threads]" << std::endl;
    std::cout << "  directory:        Directory to search for text files" << std::endl;
    std::cout << "  expressions_file: Path to expressions.properties file" << std::endl;
    std::cout << "  output_file:      Base name for output CSV files" << std::endl;
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
// First install libxlsxwriter:
//   Ubuntu/Debian: sudo apt-get install libxlsxwriter-dev
//   CentOS/RHEL: sudo yum install libxlsxwriter-devel
//   macOS: brew install libxlsxwriter
//
// Then compile:
// g++ -std=c++17 -pthread -O2 -o regex_analyzer regex_analyzer.cpp -lxlsxwriter
//
// Or with additional optimization:
// g++ -std=c++17 -pthread -O3 -march=native -o regex_analyzer regex_analyzer.cpp -lxlsxwriter
