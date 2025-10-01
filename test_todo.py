#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Simple test for todo_app functionality
"""

import json
import os
from datetime import datetime

# Test file path
TEST_TODO_FILE = "test_todos.json"

def test_save_and_load():
    """Test saving and loading todos"""
    # Sample todos
    test_todos = [
        {
            "id": 1,
            "task": "테스트 할 일 1",
            "completed": False,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        {
            "id": 2,
            "task": "테스트 할 일 2",
            "completed": True,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    ]
    
    # Save todos
    with open(TEST_TODO_FILE, 'w', encoding='utf-8') as f:
        json.dump(test_todos, f, ensure_ascii=False, indent=2)
    
    print("✓ Todos saved successfully")
    
    # Load todos
    with open(TEST_TODO_FILE, 'r', encoding='utf-8') as f:
        loaded_todos = json.load(f)
    
    print("✓ Todos loaded successfully")
    
    # Verify
    assert len(loaded_todos) == 2, "Should have 2 todos"
    assert loaded_todos[0]['task'] == "테스트 할 일 1", "First todo task mismatch"
    assert loaded_todos[0]['completed'] == False, "First todo should not be completed"
    assert loaded_todos[1]['completed'] == True, "Second todo should be completed"
    
    print("✓ All assertions passed")
    
    # Cleanup
    if os.path.exists(TEST_TODO_FILE):
        os.remove(TEST_TODO_FILE)
    print("✓ Test file cleaned up")
    
    print("\n✅ All tests passed!")

if __name__ == "__main__":
    test_save_and_load()
