"""
Compare debug upload copies with the actual file on disk
This will show if the API is working correctly
"""

from pathlib import Path
import hashlib
from loguru import logger
import os

def list_and_compare_debug_files():
    """List all debug upload files and compare with target"""
    
    debug_dir = Path("data/debug_uploads")
    
    if not debug_dir.exists():
        logger.error(f"Debug directory does not exist: {debug_dir}")
        return
    
    debug_files = sorted(debug_dir.glob("*.xlsx"), key=os.path.getmtime, reverse=True)
    
    if not debug_files:
        logger.warning("No debug upload files found")
        return
    
    logger.info("="*70)
    logger.info("DEBUG UPLOAD FILES (newest first)")
    logger.info("="*70)
    
    for debug_file in debug_files:
        logger.info(f"\nDebug file: {debug_file.name}")
        
        # Parse filename to extract file_id and hash
        parts = debug_file.stem.split('_')
        if len(parts) >= 3:
            file_id = parts[0]
            
            # Look for actual file
            actual_dir = Path("data/excel_files")
            actual_files = list(actual_dir.glob(f"{file_id}_*.xlsx"))
            
            if actual_files:
                actual_file = actual_files[0]
                
                # Read both files
                with open(debug_file, 'rb') as f:
                    debug_content = f.read()
                
                with open(actual_file, 'rb') as f:
                    actual_content = f.read()
                
                debug_hash = hashlib.sha256(debug_content).hexdigest()[:16]
                actual_hash = hashlib.sha256(actual_content).hexdigest()[:16]
                
                logger.info(f"  Debug size: {len(debug_content):,} bytes")
                logger.info(f"  Debug hash: {debug_hash}")
                logger.info(f"  Actual file: {actual_file.name}")
                logger.info(f"  Actual size: {len(actual_content):,} bytes")
                logger.info(f"  Actual hash: {actual_hash}")
                
                if debug_hash == actual_hash:
                    logger.success("  ✓ FILES MATCH - API saved correctly!")
                else:
                    logger.error("  ❌ FILES DO NOT MATCH - Something is wrong!")
                    
                    # Find first difference
                    min_len = min(len(debug_content), len(actual_content))
                    for i in range(min_len):
                        if debug_content[i] != actual_content[i]:
                            logger.error(f"  ❌ First difference at byte {i}")
                            logger.error(f"     Debug: {debug_content[max(0,i-5):i+5].hex()}")
                            logger.error(f"     Actual: {actual_content[max(0,i-5):i+5].hex()}")
                            break
            else:
                logger.warning(f"  No actual file found for file_id: {file_id}")
    
    logger.info("\n" + "="*70)
    logger.info(f"Total debug files: {len(debug_files)}")
    logger.info("="*70)
    
    # Recommend checking the most recent
    if debug_files:
        most_recent = debug_files[0]
        logger.info(f"\n💡 Most recent upload: {most_recent.name}")
        logger.info(f"   Open this file to see what the frontend sent!")
        logger.info(f"   Full path: {most_recent.resolve()}")


if __name__ == "__main__":
    logger.info("Analyzing debug upload files...\n")
    list_and_compare_debug_files()
