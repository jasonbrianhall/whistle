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
    $(info Building with XML Spreadsheet 2003 fallback (libxlsxwriter not found))
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

# Force XML mode (no XLSX even if available)
.PHONY: xml-only
xml-only: CXXFLAGS := $(filter-out -DHAVE_XLSXWRITER,$(CXXFLAGS))
xml-only: LDFLAGS := $(filter-out -lxlsxwriter,$(LDFLAGS))
xml-only: clean $(TARGET)
	@echo "Built XML Spreadsheet 2003 only version"

# Install system dependencies
.PHONY: install-deps
install-deps:
	@echo "Installing dependencies for your system..."
	@if command -v apt-get >/dev/null 2>&1; then \
		echo "Detected Debian/Ubuntu"; \
		sudo apt-get update && sudo apt-get install -y libxlsxwriter-dev || echo "libxlsxwriter not available in repos"; \
	elif command -v yum >/dev/null 2>&1; then \
		echo "Detected RHEL/CentOS"; \
		echo "libxlsxwriter not available in standard repos - will use XML Spreadsheet 2003 fallback"; \
	elif command -v dnf >/dev/null 2>&1; then \
		echo "Detected Fedora"; \
		sudo dnf install -y libxlsxwriter-devel || echo "libxlsxwriter not available in repos"; \
	elif command -v brew >/dev/null 2>&1; then \
		echo "Detected macOS"; \
		brew install libxlsxwriter || echo "libxlsxwriter not available via brew"; \
	elif command -v pacman >/dev/null 2>&1; then \
		echo "Detected Arch Linux"; \
		sudo pacman -S libxlsxwriter || echo "libxlsxwriter not available"; \
	else \
		echo "Package manager not detected."; \
		echo "libxlsxwriter not available - will use XML Spreadsheet 2003 fallback"; \
	fi
	@echo "Dependency installation attempt complete"

# Check if dependencies are available
.PHONY: check-deps
check-deps:
	@echo "Checking dependencies..."
	@if pkg-config --exists libxlsxwriter; then \
		echo "✓ libxlsxwriter found - will use modern XLSX output"; \
		pkg-config --modversion libxlsxwriter; \
	else \
		echo "⚠ libxlsxwriter not found - will use XML Spreadsheet 2003 fallback"; \
		echo "  XML Spreadsheet 2003 is fully Excel-compatible but lacks some advanced formatting"; \
		echo "  To get XLSX support, install libxlsxwriter through your package manager"; \
		echo "  or build it from source: https://libxlsxwriter.github.io/getting_started.html"; \
	fi

# Clean build artifacts
.PHONY: clean
clean:
	rm -rf $(BUILD_DIR) $(BIN_DIR)
	@echo "Cleaned build artifacts"

# Clean everything including downloaded source
.PHONY: distclean
distclean: clean
	@echo "Cleaned build artifacts"

# Display usage information
.PHONY: help
help:
	@echo "Multi-threaded Text File Regex Analyzer - Makefile"
	@echo ""
	@echo "Available targets:"
	@echo "  all                  - Build the program (default)"
	@echo "  release              - Build optimized release version"
	@echo "  debug                - Build debug version"
	@echo "  xml-only             - Build with XML Spreadsheet 2003 output only"
	@echo "  install-deps         - Install system dependencies (if available in repos)"
	@echo "  check-deps           - Check if dependencies are installed"
	@echo "  clean                - Remove build artifacts"
	@echo "  distclean            - Remove build artifacts and downloaded source"
	@echo "  help                 - Show this help message"
	@echo ""
	@echo "Output Formats:"
	@echo "  With libxlsxwriter:    Modern XLSX files with advanced formatting"
	@echo "  Without libxlsxwriter: XML Spreadsheet 2003 (Excel compatible)"
	@echo ""
	@echo "RHEL8 Usage:"
	@echo "  make install-deps       # Try to install libxlsxwriter (may not be available)"
	@echo "  make                    # Build (will auto-fallback to XML if needed)"
	@echo "  # OR"
	@echo "  make xml-only          # Build with XML Spreadsheet 2003 output"
	@echo ""
	@echo "Usage after building:"
	@echo "  ./bin/whistle <directory> <expressions_file> <output_file> [num_threads]"
	@echo ""
	@echo "Example:"
	@echo "  ./bin/whistle /var/log expressions.properties results 8"

# Show build information
.PHONY: info
info:
	@echo "Build Information:"
	@echo "  Compiler: $(CXX)"
	@echo "  Flags: $(CXXFLAGS)"
	@echo "  Libraries: $(LDFLAGS)"
	@echo "  XLSX Support: $(XLSX_AVAILABLE)"
	@echo "  Source: $(SOURCES)"
	@echo "  Target: $(TARGET)"

# Install the binary to system path
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
