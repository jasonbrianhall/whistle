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

class RegexAnalyzer {
private:
    std::vector<ExpressionPattern> expressions;
    std::vector<std::string> file_queue;
    std::mutex queue_mutex;
    std::mutex findings_mutex;
    std::vector<Finding> all_findings;
    ProgressTracker progress;
    
    std::vector<ExpressionPattern> loadExpressions(const std::string& filename);
    bool isTextFile(const std::string& filepath);
    void processFile(const std::string& filepath);
    std::vector<std::string> findTextFiles(const std::string& directory);
    void workerThread();
    
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
