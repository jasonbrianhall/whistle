#ifndef WHISTLE_H
#define WHISTLE_H

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
#include <cstring>
#include <condition_variable>
#include <memory>

// Check for libxlsxwriter availability
#ifdef HAVE_XLSXWRITER
#include <xlsxwriter.h>
#define USE_XLSX 1
#else
#define USE_XLSX 0
#endif

// XML Spreadsheet 2003 writer (fallback when XLSX not available)
class XMLSpreadsheetWriter {
private:
    std::ofstream file;
    std::map<std::string, std::vector<std::vector<std::string>>> worksheets;
    
    std::string escapeXML(const std::string& text);
    std::string cleanSheetName(const std::string& name);
    
public:
    XMLSpreadsheetWriter(const std::string& filename);
    ~XMLSpreadsheetWriter();
    
    void addWorksheet(const std::string& name);
    void addRow(const std::string& worksheet_name, const std::vector<std::string>& row);
    bool writeFile();
    bool isOpen() const;
};

struct Finding {
    std::string expression_name;
    std::string filename;
    int line_number;
    std::string actual_match;  // The actual text that matched the regex
    std::string statement;     // The full line containing the match
};

struct ExpressionPattern {
    std::string name;
    std::regex pattern;
};

// Work item representing a file-expression pair
struct WorkItem {
    std::string filepath;
    size_t expression_index;
    
    WorkItem(const std::string& path, size_t expr_idx) 
        : filepath(path), expression_index(expr_idx) {}
    
    WorkItem() : filepath(""), expression_index(0) {}
};

class Logger {
private:
    std::ofstream logFile;
    std::mutex logMutex;
    
public:
    Logger(const std::string& filename);
    ~Logger();
    
    void log(const std::string& message);
    void error(const std::string& message);
};

class ProgressTracker {
private:
    std::atomic<int> processed{0};
    std::atomic<int> total{0};
    std::chrono::steady_clock::time_point start_time;
    mutable std::mutex print_mutex;
    
public:
    void setTotal(int t);
    void increment();
    void printProgress() const;
};

class AsyncRegexAnalyzer {
private:
    std::vector<ExpressionPattern> expressions;
    
    // Thread-safe work queue
    std::queue<WorkItem> work_queue;
    std::mutex queue_mutex;
    std::condition_variable queue_cv;
    std::atomic<bool> shutdown{false};
    
    // Thread-safe findings storage
    std::mutex findings_mutex;
    std::vector<Finding> all_findings;
    
    ProgressTracker progress;
    std::unique_ptr<Logger> logger;
    
    // Thread-safe counters for debugging
    std::atomic<int> active_workers{0};
    std::atomic<int> completed_items{0};
    
    std::vector<ExpressionPattern> loadExpressions(const std::string& filename);
    bool isTextFile(const std::string& filepath);
    void processWorkItem(const WorkItem& work_item);
    std::vector<std::string> findTextFiles(const std::string& directory);
    void workerThread(int thread_id);
    
#if USE_XLSX
    void writeXLSXResults(const std::string& output_filename);
#endif
    void writeXMLSpreadsheetResults(const std::string& output_filename);
    
public:
    void analyze(const std::string& directory, const std::string& expressions_file, 
                const std::string& output_file, int num_threads = 4);
    void writeResults(const std::string& output_filename);
};

void printUsage(const char* program_name);

#endif // WHISTLE_H
