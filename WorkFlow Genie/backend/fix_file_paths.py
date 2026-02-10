"""
Utility script to fix relative file paths in database to absolute paths
Run this once to update existing records
"""

from pathlib import Path
from app.core.database import SessionLocal
from app.core.models import ExcelFile
from loguru import logger

def fix_file_paths():
    """Convert all relative file paths to absolute paths in database"""
    
    db = SessionLocal()
    try:
        # Get all excel files
        files = db.query(ExcelFile).all()
        
        logger.info(f"Found {len(files)} files in database")
        
        updated_count = 0
        for file in files:
            old_path = file.filepath
            path_obj = Path(old_path)
            
            # Check if path is already absolute
            if not path_obj.is_absolute():
                # Convert to absolute
                absolute_path = path_obj.resolve()
                
                logger.info(f"Converting path for file {file.id}:")
                logger.info(f"  Old: {old_path}")
                logger.info(f"  New: {absolute_path}")
                
                # Check if file actually exists
                if absolute_path.exists():
                    file.filepath = str(absolute_path)
                    updated_count += 1
                    logger.success(f"  ✓ Updated (file exists)")
                else:
                    logger.warning(f"  ✗ File not found at resolved path: {absolute_path}")
            else:
                logger.info(f"File {file.id} already has absolute path: {old_path}")
        
        # Commit changes
        if updated_count > 0:
            db.commit()
            logger.success(f"Successfully updated {updated_count} file paths")
        else:
            logger.info("No paths needed updating")
            
    except Exception as e:
        logger.error(f"Error fixing paths: {e}")
        db.rollback()
        raise
    finally:
        db.close()

if __name__ == "__main__":
    logger.info("Starting file path fix utility...")
    fix_file_paths()
    logger.info("Done!")
