"""
Debug script to check file status, locks, and permissions
Run this to diagnose file save issues
"""

from pathlib import Path
import os
import stat
from loguru import logger

def check_file_status(file_path: str):
    """Check and report file status, permissions, and locks"""
    
    path = Path(file_path).resolve()
    
    logger.info("="*70)
    logger.info(f"Checking file: {path}")
    logger.info("="*70)
    
    # Check if file exists
    if not path.exists():
        logger.error(f"❌ File does NOT exist at: {path}")
        return
    
    logger.success(f"✓ File exists at: {path}")
    
    # Check file size
    size = path.stat().st_size
    logger.info(f"File size: {size:,} bytes ({size / 1024:.2f} KB)")
    
    # Check file permissions
    file_stat = path.stat()
    mode = file_stat.st_mode
    
    logger.info("\nFile Permissions:")
    logger.info(f"  Owner can read: {bool(mode & stat.S_IRUSR)}")
    logger.info(f"  Owner can write: {bool(mode & stat.S_IWUSR)}")
    logger.info(f"  Owner can execute: {bool(mode & stat.S_IXUSR)}")
    
    # Check if file is read-only
    is_readonly = not (mode & stat.S_IWUSR)
    if is_readonly:
        logger.warning(f"⚠️  File is READ-ONLY!")
    else:
        logger.success(f"✓ File is writable")
    
    # Check if file is locked (Windows)
    if os.name == 'nt':
        logger.info("\nChecking file lock status (Windows)...")
        try:
            import msvcrt
            with open(path, 'r+b') as f:
                try:
                    msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                    logger.success("✓ File is NOT locked")
                except IOError:
                    logger.error("❌ File IS LOCKED (probably open in Excel or another program)")
        except Exception as e:
            logger.error(f"❌ Cannot open file for lock check: {e}")
            logger.warning("  This usually means the file is in use by another program")
    
    # Try to read file
    logger.info("\nTrying to read file...")
    try:
        with open(path, 'rb') as f:
            data = f.read()
        logger.success(f"✓ Successfully read {len(data)} bytes")
    except Exception as e:
        logger.error(f"❌ Cannot read file: {e}")
    
    # Try to open file for writing
    logger.info("\nTrying to open file for writing...")
    try:
        with open(path, 'r+b') as f:
            pass
        logger.success("✓ File can be opened for writing")
    except Exception as e:
        logger.error(f"❌ Cannot open file for writing: {e}")
    
    # Check parent directory permissions
    logger.info("\nParent Directory:")
    parent = path.parent
    logger.info(f"  Directory: {parent}")
    logger.info(f"  Exists: {parent.exists()}")
    
    if parent.exists():
        parent_stat = parent.stat()
        parent_mode = parent_stat.st_mode
        logger.info(f"  Can write to directory: {bool(parent_mode & stat.S_IWUSR)}")
    
    logger.info("="*70)
    logger.info("Diagnosis complete")
    logger.info("="*70)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python debug_file_lock.py <file_path>")
        print("\nExample:")
        print('python debug_file_lock.py "data\\excel_files\\100f49aa-3ad9-4115-b209-503f9da44c89_TrialBal.xlsx"')
        sys.exit(1)
    
    file_path = sys.argv[1]
    check_file_status(file_path)
