try:
    import google.generativeai
    print("SUCCESS: google.generativeai is installed.")
except ImportError:
    print("FAILURE: google.generativeai is NOT installed.")
    import sys
    sys.exit(1)
