# Makefile for Multi-threaded Text File Regex Analyzer

# Compiler and flags
CXX = g++
CXXFLAGS = -std=c++17 -pthread -Wall -Wextra -O2
LDFLAGS = 

# Check for libxlsxwriter availability
XLSX_AVAILABLE := $(shell pkg-config --exists libxlsxwriter && echo 1 || echo 0)

ifeq ($(XLSX_AVAILABLE),1)
    CXXFLAGS += -DHAVE_XLSXWRITER
    LDFLAGS += -lxlsxwriter
    $(info Building with XLSX support)
else
    $(info Building with CSV fallback (libxlsxwriter not found))
endif

# Optimized build flags
OPT_FLAGS = -O3 -march=native -DNDEBUG

# Debug flags
DEBUG_FLAGS = -g -O0 -DDEBUG

# Directories
SRC_DIR = .
BUILD_DIR = build
BIN_DIR = bin

# Source files
SOURCES = whistle.cpp
OBJECTS = $(BUILD_DIR)/whistle.o
TARGET = $(BIN_DIR)/whistle

# Default target
.PHONY: all
all: $(TARGET)

# Create directories if they don't exist
$(BUILD_DIR):
	mkdir -p $(BUILD_DIR)

$(BIN_DIR):
	mkdir -p $(BIN_DIR)

# Compile object files
$(BUILD_DIR)/%.o: $(SRC_DIR)/%.cpp | $(BUILD_DIR)
	$(CXX) $(CXXFLAGS) -c $< -o $@

# Link executable
$(TARGET): $(OBJECTS) | $(BIN_DIR)
	$(CXX) $(OBJECTS) $(LDFLAGS) -o $@

# Optimized build
.PHONY: release
release: CXXFLAGS += $(OPT_FLAGS)
release: clean $(TARGET)
	@echo "Built optimized release version"

# Debug build
.PHONY: debug
debug: CXXFLAGS += $(DEBUG_FLAGS)
debug: clean $(TARGET)
	@echo "Built debug version"

# Force CSV mode (no XLSX even if available)
.PHONY: csv-only
csv-only: CXXFLAGS := $(filter-out -DHAVE_XLSXWRITER,$(CXXFLAGS))
csv-only: LDFLAGS := $(filter-out -lxlsxwriter,$(LDFLAGS))
csv-only: clean $(TARGET)
	@echo "Built CSV-only version"

# Install system dependencies
.PHONY: install-deps
install-deps:
	@echo "Installing dependencies for your system..."
	@if command -v apt-get >/dev/null 2>&1; then \
		echo "Detected Debian/Ubuntu"; \
		sudo apt-get update && sudo apt-get install -y libxlsxwriter-dev; \
	elif command -v yum >/dev/null 2>&1; then \
		echo "Detected RHEL/CentOS - building from source"; \
		$(MAKE) install-xlsx-from-source; \
	elif command -v dnf >/dev/null 2>&1; then \
		echo "Detected Fedora"; \
		sudo dnf install -y libxlsxwriter-devel || $(MAKE) install-xlsx-from-source; \
	elif command -v brew >/dev/null 2>&1; then \
		echo "Detected macOS"; \
		brew install libxlsxwriter; \
	elif command -v pacman >/dev/null 2>&1; then \
		echo "Detected Arch Linux"; \
		sudo pacman -S libxlsxwriter; \
	else \
		echo "Package manager not detected. Building from source..."; \
		$(MAKE) install-xlsx-from-source; \
	fi
	@echo "Dependencies installation complete"

# Build libxlsxwriter from source (for RHEL8, etc.)
.PHONY: install-xlsx-from-source
install-xlsx-from-source:
	@echo "Building libxlsxwriter from source..."
	@if [ ! -d "libxlsxwriter-build" ]; then \
		echo "Downloading libxlsxwriter..."; \
		wget -O libxlsxwriter.tar.gz# Makefile for Multi-threaded Text File Regex Analyzer

# Compiler and flags
CXX = g++
CXXFLAGS = -std=c++17 -pthread -Wall -Wextra -O2
LDFLAGS = -lxlsxwriter

# Optimized build flags
OPT_FLAGS = -O3 -march=native -DNDEBUG

# Debug flags
DEBUG_FLAGS = -g -O0 -DDEBUG

# Directories
SRC_DIR = .
BUILD_DIR = build
BIN_DIR = bin

# Source files
SOURCES = whistle.cpp
OBJECTS = $(BUILD_DIR)/whistle.o
TARGET = $(BIN_DIR)/whistle

# Default target
.PHONY: all
all: $(TARGET)

# Create directories if they don't exist
$(BUILD_DIR):
	mkdir -p $(BUILD_DIR)

$(BIN_DIR):
	mkdir -p $(BIN_DIR)

# Compile object files
$(BUILD_DIR)/%.o: $(SRC_DIR)/%.cpp | $(BUILD_DIR)
	$(CXX) $(CXXFLAGS) -c $< -o $@

# Link executable
$(TARGET): $(OBJECTS) | $(BIN_DIR)
	$(CXX) $(OBJECTS) $(LDFLAGS) -o $@

# Optimized build
.PHONY: release
release: CXXFLAGS += $(OPT_FLAGS)
release: clean $(TARGET)
	@echo "Built optimized release version"

# Debug build
.PHONY: debug
debug: CXXFLAGS += $(DEBUG_FLAGS)
debug: clean $(TARGET)
	@echo "Built debug version"

# Install system dependencies
.PHONY: install-deps
install-deps:
	@echo "Installing libxlsxwriter..."
	@if command -v apt-get >/dev/null 2>&1; then \
		sudo apt-get update && sudo apt-get install -y libxlsxwriter-dev; \
	elif command -v yum >/dev/null 2>&1; then \
		sudo yum install -y libxlsxwriter-devel; \
	elif command -v dnf >/dev/null 2>&1; then \
		sudo dnf install -y libxlsxwriter-devel; \
	elif command -v brew >/dev/null 2>&1; then \
		brew install libxlsxwriter; \
	elif command -v pacman >/dev/null 2>&1; then \
		sudo pacman -S libxlsxwriter; \
	else \
		echo "Package manager not detected. Please install libxlsxwriter manually."; \
		echo "See: https://libxlsxwriter.github.io/getting_started.html"; \
		exit 1; \
	fi
	@echo "Dependencies installed successfully"

# Check if dependencies are available
.PHONY: check-deps
check-deps:
	@echo "Checking dependencies..."
	@if pkg-config --exists libxlsxwriter; then \
		echo "✓ libxlsxwriter found"; \
	else \
		echo "✗ libxlsxwriter not found"; \
		echo "Run 'make install-deps' to install dependencies"; \
		exit 1; \
	fi

# Clean build artifacts
.PHONY: clean
clean:
	rm -rf $(BUILD_DIR) $(BIN_DIR)
	@echo "Cleaned build artifacts"

# Test the program (requires test files)
.PHONY: test
test: $(TARGET)
	@echo "Running basic test..."
	@if [ ! -d "test_data" ]; then \
		mkdir -p test_data; \
		echo "https://example.com" > test_data/sample1.txt; \
		echo "192.168.1.1" > test_data/sample2.txt; \
		echo "test@email.com" > test_data/sample3.txt; \
	fi
	@if [ ! -f "test_expressions.properties" ]; then \
		echo "[expressions]" > test_expressions.properties; \
		echo "expression.url=https?://[\\w.-]+[\\w/]+" >> test_expressions.properties; \
		echo "expression.ip=\\b(?:[0-9]{1,3}\\.){3}[0-9]{1,3}\\b" >> test_expressions.properties; \
		echo "expression.email=\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b" >> test_expressions.properties; \
	fi
	$(TARGET) test_data test_expressions.properties test_output.xlsx 2
	@echo "Test completed. Check test_output.xlsx"

# Create sample configuration and test data
.PHONY: sample
sample:
	@echo "Creating sample files..."
	@mkdir -p sample_data
	@echo "[expressions]" > expressions.properties
	@echo "expression.url=https?://[\\w.-]+[\\w/]+" >> expressions.properties
	@echo "expression.ip=\\b(?:[0-9]{1,3}\\.){3}[0-9]{1,3}\\b" >> expressions.properties
	@echo "expression.email=\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b" >> expressions.properties
	@echo "expression.phone=\\b\\d{3}-\\d{3}-\\d{4}\\b" >> expressions.properties
	@echo "expression.ssn=\\b\\d{3}-\\d{2}-\\d{4}\\b" >> expressions.properties
	@echo "Visit our website at https://example.com for more info" > sample_data/website.txt
	@echo "Server IP: 192.168.1.100" > sample_data/config.log
	@echo "Contact us at support@company.com or call 555-123-4567" > sample_data/contact.txt
	@echo "SSN: 123-45-6789" > sample_data/sensitive.log
	@echo "API endpoint: https://api.service.com/v1/data" > sample_data/api.txt
	@echo "Database: 10.0.0.50:3306" > sample_data/database.conf
	@echo "Created sample files:"
	@echo "  - expressions.properties (regex patterns)"
	@echo "  - sample_data/ (test text files)"
	@echo "Run: make run-sample"

# Run with sample data
.PHONY: run-sample
run-sample: $(TARGET) sample
	$(TARGET) sample_data expressions.properties results.xlsx 4

# Display usage information
.PHONY: help
help:
	@echo "Multi-threaded Text File Regex Analyzer - Makefile"
	@echo ""
	@echo "Available targets:"
	@echo "  all          - Build the program (default)"
	@echo "  release      - Build optimized release version"
	@echo "  debug        - Build debug version"
	@echo "  install-deps - Install system dependencies (libxlsxwriter)"
	@echo "  check-deps   - Check if dependencies are installed"
	@echo "  clean        - Remove build artifacts"
	@echo "  test         - Run basic test with sample data"
	@echo "  sample       - Create sample configuration and test files"
	@echo "  run-sample   - Build and run with sample data"
	@echo "  help         - Show this help message"
	@echo ""
	@echo "Usage after building:"
	@echo "  ./bin/whistle <directory> <expressions_file> <output_file> [num_threads]"
	@echo ""
	@echo "Example:"
	@echo "  ./bin/whistle /path/to/files expressions.properties results.xlsx 8"

# Show build information
.PHONY: info
info:
	@echo "Build Information:"
	@echo "  Compiler: $(CXX)"
	@echo "  Flags: $(CXXFLAGS)"
	@echo "  Libraries: $(LDFLAGS)"
	@echo "  Source: $(SOURCES)"
	@echo "  Target: $(TARGET)"

# Install the binary to system path (optional)
.PHONY: install
install: $(TARGET)
	@echo "Installing whistle to /usr/local/bin..."
	sudo cp $(TARGET) /usr/local/bin/
	sudo chmod +x /usr/local/bin/whistle
	@echo "Installation complete. You can now run 'whistle' from anywhere."

# Uninstall from system
.PHONY: uninstall
uninstall:
	@echo "Removing whistle from /usr/local/bin..."
	sudo rm -f /usr/local/bin/whistle
	@echo "Uninstall complete."

# Force rebuild
.PHONY: rebuild
rebuild: clean all

# Show file sizes and build stats
.PHONY: stats
stats: $(TARGET)
	@echo "Build Statistics:"
	@echo "  Executable size: $$(du -h $(TARGET) | cut -f1)"
	@echo "  Object files: $$(find $(BUILD_DIR) -name '*.o' | wc -l)"
	@echo "  Build time: $$(stat -c %y $(TARGET) 2>/dev/null || stat -f %Sm $(TARGET) 2>/dev/null || echo 'Unknown')"
