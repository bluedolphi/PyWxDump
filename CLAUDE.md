# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PyWxDump is a Python tool for extracting WeChat account information, decrypting WeChat databases, and viewing/exporting chat history. The project provides functionality to obtain WeChat user data (nicknames, accounts, phones, emails, database keys), decrypt encrypted databases, and export conversations in various formats.

**This is a security research tool designed for legitimate purposes like data backup and forensic analysis.**

## Installation and Common Commands

### Installation
```bash
# Install from PyPI
pip install pywxdump

# Install from source
git clone https://github.com/xaoyaoo/PyWxDump.git
cd PyWxDump
pip install -r requirements.txt
python setup.py install

# Install in development mode
pip install -e .
```

### CLI Usage
```bash
# Main CLI entry point
wxdump --help

# Get WeChat information (requires WeChat to be running)
wxdump info

# Get bias address for memory reading
wxdump bias

# Decrypt databases
wxdump decrypt --key <key> --db <db_path> --out <output_path>

# Merge databases
wxdump merge --db <db_path> --out <output_path>

# Start web server for viewing chat history
wxdump server --port <port>

# Export chat history
wxdump export --html <output_path>
wxdump export --csv <output_path>
wxdump export --json <output_path>
```

### Development Commands
```bash
# Run tests
python -m pytest tests/

# Build executable
python tests/build_exe.py

# Generate changelog
python tests/gen_change_log.py

# Release new version
python tests/release_new_version.py

# Test specific functionality
python tests/test_decrypt.py
python tests/test_read_info.py
python tests/test_Bias.py
```

## Architecture Overview

### Core Modules

#### `pywxdump/wx_core/` - WeChat Core Functionality
- **`wx_info.py`**: Extracts WeChat account information from running process memory
- **`get_bias_addr.py`**: Memory address calculation for different WeChat versions
- **`decryption.py`**: Database decryption using AES and other algorithms
- **`memory_search.py`**: Low-level memory scanning utilities
- **`merge_db.py`**: Database merging and combination utilities

#### `pywxdump/db/` - Database Handlers
- **`dbbase.py`**: Base database connection pooling and management
- **`dbMSG.py`**: Message/chat database operations
- **`dbMedia.py`**: Media file database operations
- **`dbMicro.py`**: Micro-message database operations
- **`dbFavorite.py`**: Favorites database operations
- **`dbPublicMsg.py`**: Public message database operations

#### `pywxdump/api/` - API and Server Components
- **`local_server.py`**: FastAPI-based local web server for chat viewing
- **`remote_server.py`**: Remote server functionality
- **`export/`**: Export modules for different formats (HTML, CSV, JSON)

#### `pywxdump/analyzer/` - Data Analysis
- **`chat_analysis.py`**: Chat history analysis and statistics
- **`cleanup.py`**: Data cleanup and optimization utilities

### Key Configuration Files

#### `WX_OFFS.json`
Critical configuration file containing memory address offsets for different WeChat versions. Maps WeChat version numbers to memory addresses used for extracting account information.

**Note**: This file is regularly updated as WeChat versions change. Contributions of new offsets are welcome.

### Data Flow Architecture

1. **Memory Extraction**: `wx_core/wx_info.py` reads WeChat process memory
2. **Address Calculation**: `get_bias_addr.py` calculates version-specific offsets
3. **Database Decryption**: `decryption.py` decrypts SQLite databases using extracted keys
4. **Data Processing**: Database handlers parse decrypted data
5. **Export/Viewing**: API and export modules provide user interfaces

## Development Notes

### Dependencies
- **Memory Reading**: `pymem`, `psutil`, `pywin32` (Windows-specific)
- **Cryptography**: `pycryptodomex` for database decryption
- **Web Framework**: `FastAPI`, `uvicorn` for web interface
- **Database**: `sqlite3`, `dbutils` for connection pooling
- **Media Processing**: `silk-python`, `pyaudio` for voice message handling

### Platform Support
- **Primary**: Windows (WeChat desktop client)
- **Experimental**: macOS and Linux support may have limitations
- **Python Version**: Requires Python 3.8+

### Testing Approach
- Tests are located in `tests/` directory
- Test files require actual WeChat data paths and keys
- Use placeholder values in test files and replace with real data for testing

### Security Considerations
- Only works on locally installed WeChat instances
- Requires WeChat to be running for memory extraction
- All operations are read-only (no modification of WeChat data)
- Designed for legitimate backup and forensic purposes

## Important Files for Development

- `pywxdump/cli.py`: Main CLI interface and argument parsing
- `pywxdump/__init__.py`: Module exports and version management
- `pywxdump/WX_OFFS.json`: Version-specific memory offsets (critical for functionality)
- `requirements.txt`: Python dependencies
- `setup.py`: Package configuration and installation

## Common Development Tasks

### Adding Support for New WeChat Versions
1. Use Cheat Engine to find new memory addresses (see `doc/CE获取基址.md`)
2. Add version entries to `WX_OFFS.json`
3. Test with `wxdump bias` command
4. Submit PR with version information

### Database Schema Changes
- Modify relevant handlers in `pywxdump/db/`
- Update export modules if needed
- Test with sample databases

### New Export Formats
- Add new module in `pywxdump/api/export/`
- Register in `__init__.py` exports
- Update CLI options in `cli.py`