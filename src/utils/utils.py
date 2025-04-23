import sys
import os

def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def is_file_open(file_path):
        """Check if file is open"""
        try:
            with open(file_path, "r+"):
                pass
        except IOError:
            return True  
        return False  

