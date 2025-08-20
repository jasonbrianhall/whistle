# Makefile for Multi-threaded Text File Regex Analyzer

# Compiler and flags
CXX = g++
CXXFLAGS = -std=c++17 -pthread -Wall -Wextra -O2
LDFLAGS = 

# Force XML mode flag
XML_ONLY = 0

# Set flags based on mode
ifeq ($(XML_ONLY),0)
    # Try XLSX first
    CXXFLAGS += -DHAVE_XLSXWRITER
    XLSX_LIBS = -lxlsxwriter -lpthread
    $(info Attempting to build with XLSX support)
else
    XLSX_LIBS = -lpthread
    $(info Building with XML Spreadsheet 2003 output only (xml-only specified))
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

# Default target - try XLSX first, fallback to XML if it fails
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

# Link executable with fallback logic
$(TARGET): $(OBJECTS) | $(BIN_DIR)
ifeq ($(XML_ONLY),0)
	@echo "Linking with libxlsxwriter..."
	@$(CXX) $(OBJECTS) $(XLSX_LIBS) -o $@ 2>/dev/null || \
	(echo "libxlsxwriter linking failed - rebuilding with XML fallback..." && \
	 $(MAKE) xml-only-internal)
else
	$(CXX) $(OBJECTS) $(XLSX_LIBS) -o $@
endif

# Internal target for automatic fallback (don't clean, just rebuild)
.PHONY: xml-only-internal
xml-only-internal:
	@echo "Rebuilding object files without XLSX support..."
	@$(CXX) -std=c++17 -pthread -Wall -Wextra -O2 -c $(SRC_DIR)/whistle.cpp -o $(BUILD_DIR)/whistle.o
	@$(CXX) $(BUILD_DIR)/whistle.o -lpthread -o $(TARGET)
	@echo "Successfully built with XML Spreadsheet 2003 fallback"

# Optimized build
.PHONY: release
release: CXXFLAGS += $(OPT_FLAGS)
release: clean
	@$(MAKE) all XLSX_LIBS="$(XLSX_LIBS)"
	@echo "Built optimized release version"

# Debug build
.PHONY: debug
debug: CXXFLAGS += $(DEBUG_FLAGS)
debug: clean
	@$(MAKE) all XLSX_LIBS="$(XLSX_LIBS)"
	@echo "Built debug version"

# Force XML mode (no XLSX even if available)
.PHONY: xml-only
xml-only: XML_ONLY = 1
xml-only: clean
	@$(MAKE) all XML_ONLY=1 XLSX_LIBS="-lpthread"
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
	@echo "Note: This Makefile will automatically try libxlsxwriter and fallback to XML if needed"
	@if command -v $(CXX) >/dev/null 2>&1; then \
		echo "✓ C++ compiler found: $(CXX)"; \
	else \
		echo "✗ C++ compiler not found"; \
	fi
	@echo "Build will attempt XLSX first, then fallback to XML Spreadsheet 2003 if linking fails"

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
