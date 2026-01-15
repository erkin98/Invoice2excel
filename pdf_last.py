#!/usr/bin/env python
# coding: utf-8

import sys
import os

# Ensure current directory is in path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from main import main

if __name__ == "__main__":
    print("Starting Invoice2Excel...")
    main()
