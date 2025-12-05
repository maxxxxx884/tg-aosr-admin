#!/usr/bin/env python3
"""
Test script to verify axes and marks functionality in OZR and ZVK modules
"""

import json
import os
import tempfile
import shutil
from OZR import ProductionJournalEditor
from ZVK import IncomingJournalEditor
import tkinter as tk


def test_production_journal_axes_marks():
    """Test that axes and marks fields are properly handled in production journal"""
    print("Testing Production Journal Axes and Marks functionality...")
    
    # Create a temporary directory structure for testing
    temp_dir = tempfile.mkdtemp()
    contractor_dir = os.path.join(temp_dir, "Test Contractor")
    journal_dir = os.path.join(contractor_dir, "Журнал производства работ")
    os.makedirs(journal_dir)
    
    # Create a test journal_production.json file
    test_data = {
        "entries": [
            {
                "date": "2025-09-04",
                "name": "Штукатурка стен",
                "axes": "1",
                "marks": "+1",
                "volume": 10.0,
                "volume_unit": "м²",
                "photos": [],
                "filled_by": "Тестовый пользователь",
                "created_at": "2025-09-04 14:14:47"
            }
        ]
    }
    
    journal_file = os.path.join(journal_dir, "journal_production.json")
    with open(journal_file, 'w', encoding='utf-8') as f:
        json.dump(test_data, f, ensure_ascii=False, indent=2)
    
    # Test that the data loads correctly with axes and marks
    root = tk.Tk()
    app = ProductionJournalEditor(root)
    app.current_directory = temp_dir
    app.load_data()
    
    # Check that we have the correct number of entries
    assert len(app.data) == 1, f"Expected 1 entry, got {len(app.data)}"
    
    # Check that axes and marks are properly loaded
    entry = app.data[0]
    assert entry.get('axes') == "1", f"Expected axes '1', got '{entry.get('axes')}'"
    assert entry.get('marks') == "+1", f"Expected marks '+1', got '{entry.get('marks')}'"
    
    print("✓ Production Journal Axes and Marks loading test passed")
    
    # Clean up
    root.destroy()
    shutil.rmtree(temp_dir)


def test_incoming_journal_axes_marks():
    """Test that axes and marks fields are properly handled in incoming journal"""
    print("Testing Incoming Journal Axes and Marks functionality...")
    
    # Create a temporary directory structure for testing
    temp_dir = tempfile.mkdtemp()
    contractor_dir = os.path.join(temp_dir, "Test Contractor")
    journal_dir = os.path.join(contractor_dir, "Журнал входного контроля")
    os.makedirs(journal_dir)
    
    # Create a test journal_incoming.json file
    test_data = {
        "entries": [
            {
                "date": "2025-09-04",
                "name": "Тестовый материал",
                "axes": "1-1",
                "marks": "+0.000",
                "quantity": 10.0,
                "quantity_unit": "м²",
                "supplier": "Тестовый поставщик",
                "document": "Паспорт 123456",
                "lab_control_needed": False,
                "lab_control_result": "",
                "filled_by": "Тестовый пользователь",
                "created_at": "2025-09-04 14:14:47"
            }
        ]
    }
    
    journal_file = os.path.join(journal_dir, "journal_incoming.json")
    with open(journal_file, 'w', encoding='utf-8') as f:
        json.dump(test_data, f, ensure_ascii=False, indent=2)
    
    # Test that the data loads correctly with axes and marks
    root = tk.Tk()
    app = IncomingJournalEditor(root)
    app.current_directory = temp_dir
    app.load_data()
    
    # Check that we have the correct number of entries
    assert len(app.data) == 1, f"Expected 1 entry, got {len(app.data)}"
    
    # Check that axes and marks are properly loaded
    entry = app.data[0]
    assert entry.get('axes') == "1-1", f"Expected axes '1-1', got '{entry.get('axes')}'"
    assert entry.get('marks') == "+0.000", f"Expected marks '+0.000', got '{entry.get('marks')}'"
    
    print("✓ Incoming Journal Axes and Marks loading test passed")
    
    # Clean up
    root.destroy()
    shutil.rmtree(temp_dir)


if __name__ == "__main__":
    try:
        test_production_journal_axes_marks()
        test_incoming_journal_axes_marks()
        print("\n🎉 All tests passed! Axes and Marks functionality is working correctly.")
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        raise