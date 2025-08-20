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

// Simple CSV writer for Excel compatibility
class CSVWriter {
private:
    std::ofstream file;
    
public:
    CSVWriter(const std::string& filename) : file(filename) {}
    
    void writeRow(const std::vector<std::string>& row) {
        for (size_t i = 0; i < row.size(); ++i) {
            if (i > 0) file << ",";
            // Escape quotes and wrap in quotes if contains comma or quote
            std::string cell = row[i];
            if (cell.find(',') != std::string::npos || cell.find('"') != std::string::npos) {
                // Escape existing quotes by doubling them
                size_t pos = 0;
                while ((pos = cell.find('"', pos)) != std::string::npos) {
                    cell.replace(pos, 1, "\"\"");
                    pos += 2;
                }
                file << "\"" << cell << "\"";
            } else {
                file << cell;
            }
        }
        file << std::endl;
    }
    
    ~CSVWriter() {
        if (file.is_open()) {
            file.close();
        }
    }
};

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
    void writeResults(const std::string& base_filename) {
        // Group findings by expression
        std::map<std::string, std::vector<Finding>> grouped_findings;
        
        for (const auto& finding : all_findings) {
            grouped_findings[finding.expression_name].push_back(finding);
        }
        
        // Create separate CSV files for each expression (Excel tabs)
        for (const auto& [expr_name, findings] : grouped_findings) {
            std::string filename = base_filename;
            size_t dot_pos = filename.find_last_of('.');
            if (dot_pos != std::string::npos) {
                filename = filename.substr(0, dot_pos) + "_" + expr_name + ".csv";
            } else {
                filename += "_" + expr_name + ".csv";
            }
            
            CSVWriter writer(filename);
            
            // Write header
            writer.writeRow({"Finding", "File", "Line", "Comments", "Ease", "Significance", "Risk", "Statement"});
            
            // Write findings
            for (const auto& finding : findings) {
                writer.writeRow({
                    expr_name,
                    finding.filename,
                    std::to_string(finding.line_number),
                    "", // Comments (blank)
                    "", // Ease (blank)  
                    "", // Significance (blank)
                    "", // Risk (blank)
                    finding.statement
                });
            }
            
            std::cout << "Created: " << filename << " with " << findings.size() << " findings" << std::endl;
        }
        
        // Also create a summary file with all findings
        std::string summary_filename = base_filename;
        size_t dot_pos = summary_filename.find_last_of('.');
        if (dot_pos != std::string::npos) {
            summary_filename = summary_filename.substr(0, dot_pos) + "_summary.csv";
        } else {
            summary_filename += "_summary.csv";
        }
        
        CSVWriter summary_writer(summary_filename);
        summary_writer.writeRow({"Finding", "File", "Line", "Comments", "Ease", "Significance", "Risk", "Statement"});
        
        for (const auto& finding : all_findings) {
            summary_writer.writeRow({
                finding.expression_name,
                finding.filename,
                std::to_string(finding.line_number),
                "", "", "", "",
                finding.statement
            });
        }
        
        std::cout << "Created summary: " << summary_filename << " with " << all_findings.size() << " total findings" << std::endl;
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
// g++ -std=c++17 -pthread -O2 -o regex_analyzer regex_analyzer.cpp
//
// Or with additional optimization:
// g++ -std=c++17 -pthread -O3 -march=native -o regex_analyzer regex_analyzer.cpp
