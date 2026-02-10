"""
Compare file content with backup to prove content changed
"""

from pathlib import Path
import hashlib
from loguru import logger

def compare_files(file_path: str):
    """Compare current file with its backup"""
    
    path = Path(file_path).resolve()
    backup_path = path.with_suffix(path.suffix + '.backup')
    
    logger.info("="*70)
    logger.info("FILE COMPARISON")
    logger.info("="*70)
    
    # Check main file
    if not path.exists():
        logger.error(f"Main file does not exist: {path}")
        return
    
    with open(path, 'rb') as f:
        current_content = f.read()
    
    current_hash = hashlib.sha256(current_content).hexdigest()
    current_size = len(current_content)
    
    logger.info(f"\nCurrent File: {path.name}")
    logger.info(f"  Size: {current_size:,} bytes")
    logger.info(f"  SHA256: {current_hash}")
    logger.info(f"  Modified: {path.stat().st_mtime}")
    logger.info(f"  First 32 bytes (hex): {current_content[:32].hex()}")
    logger.info(f"  Last 32 bytes (hex): {current_content[-32:].hex()}")
    
    # Check backup if exists
    if backup_path.exists():
        with open(backup_path, 'rb') as f:
            backup_content = f.read()
        
        backup_hash = hashlib.sha256(backup_content).hexdigest()
        backup_size = len(backup_content)
        
        logger.info(f"\nBackup File: {backup_path.name}")
        logger.info(f"  Size: {backup_size:,} bytes")
        logger.info(f"  SHA256: {backup_hash}")
        logger.info(f"  First 32 bytes (hex): {backup_content[:32].hex()}")
        logger.info(f"  Last 32 bytes (hex): {backup_content[-32:].hex()}")
        
        logger.info("\n" + "="*70)
        logger.info("COMPARISON RESULTS")
        logger.info("="*70)
        
        if current_hash == backup_hash:
            logger.warning("⚠️  Files are IDENTICAL (same hash)")
        else:
            logger.success("✓ Files are DIFFERENT")
            
            if current_size == backup_size:
                logger.info(f"  Size: Same ({current_size} bytes)")
            else:
                logger.info(f"  Size: Different (current: {current_size}, backup: {backup_size})")
            
            # Find first difference
            for i, (a, b) in enumerate(zip(current_content, backup_content)):
                if a != b:
                    logger.info(f"  First difference at byte {i}: current={a:02x}, backup={b:02x}")
                    logger.info(f"  Context: ...{current_content[max(0,i-10):i+10].hex()}...")
                    break
    else:
        logger.info(f"\nNo backup file found at: {backup_path}")
    
    logger.info("\n" + "="*70)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python compare_file_hashes.py <file_path>")
        print("\nExample:")
        print('python compare_file_hashes.py "data\\excel_files\\100f49aa-3ad9-4115-b209-503f9da44c89_TrialBal.xlsx"')
        sys.exit(1)
    
    file_path = sys.argv[1]
    compare_files(file_path)
