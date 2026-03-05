#!/usr/bin/env python
# minimal_tts_update_script.py - Minimal version to identify execution issues

import os
import sys

def main():
    """
    Simplified main function to test execution
    """
    print("=== TTS Card Updater Script ===")
    print(f"Python version: {sys.version}")
    print(f"Current working directory: {os.getcwd()}")
    
    # Look for a TTS save file (.json) in the current directory
    try:
        json_files = [f for f in os.listdir(".") if f.endswith('.json') and not f.endswith('.backup') and not f.endswith('_modified.json')]
        
        if not json_files:
            print("No JSON files found in the current directory.")
        else:
            print("Found JSON files:")
            for i, f in enumerate(json_files):
                print(f"{i+1}. {f}")
            
            selection = input("\nEnter the number of the TTS save file to process: ")
            print(f"You selected: {selection}")
    except Exception as e:
        print(f"Error finding JSON files: {e}")
    
    print("\nScript execution complete!")

if __name__ == "__main__":
    try:
        print("Script starting...")
        main()
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\nExecution finished.")
        input("Press Enter to exit...")