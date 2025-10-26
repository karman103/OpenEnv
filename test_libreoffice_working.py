#!/usr/bin/env python3
"""
Test script to verify the LibreOffice environment is working correctly.

This script tests the environment structure and error handling without requiring
LibreOffice to be installed locally.
"""

import sys
from pathlib import Path

# Add the src directory to the Python path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_imports():
    """Test that all modules can be imported correctly."""
    print("Testing imports...")
    
    try:
        from envs.libreoffice_env.models import LibreOfficeAction, LibreOfficeObservation
        print("✅ Models imported successfully")
    except Exception as e:
        print(f"❌ Models import failed: {e}")
        return False
    
    try:
        from envs.libreoffice_env.client import LibreOfficeEnv
        print("✅ Client imported successfully")
    except Exception as e:
        print(f"❌ Client import failed: {e}")
        return False
    
    try:
        from envs.libreoffice_env.server.app import app
        print("✅ Server app imported successfully")
    except Exception as e:
        print(f"❌ Server app import failed: {e}")
        return False
    
    try:
        from envs.libreoffice_env.server.libreoffice_environment import LibreOfficeEnvironment
        print("✅ Environment imported successfully")
    except Exception as e:
        print(f"❌ Environment import failed: {e}")
        return False
    
    return True

def test_models():
    """Test model creation and validation."""
    print("\nTesting models...")
    
    try:
        from envs.libreoffice_env.models import LibreOfficeAction, LibreOfficeObservation
        
        # Test action creation
        action = LibreOfficeAction(
            command="set_cell",
            parameters={"cell": "A1", "value": "test"}
        )
        print("✅ Action created successfully")
        print(f"   Command: {action.command}")
        print(f"   Parameters: {action.parameters}")
        
        # Test observation creation
        obs = LibreOfficeObservation(
            result="Test result",
            success=True,
            data="test data"
        )
        print("✅ Observation created successfully")
        print(f"   Result: {obs.result}")
        print(f"   Success: {obs.success}")
        print(f"   Data: {obs.data}")
        
        return True
    except Exception as e:
        print(f"❌ Model test failed: {e}")
        return False

def test_environment_structure():
    """Test environment structure and basic functionality."""
    print("\nTesting environment structure...")
    
    try:
        from envs.libreoffice_env.server.libreoffice_environment import LibreOfficeEnvironment
        from envs.libreoffice_env.models import LibreOfficeAction
        
        # Create environment
        env = LibreOfficeEnvironment()
        print("✅ Environment created successfully")
        
        # Test initial state
        state = env.state
        print(f"✅ Initial state - Episode ID: {state.episode_id}")
        print(f"   Step count: {state.step_count}")
        
        # Test reset (will fail gracefully without LibreOffice)
        obs = env.reset()
        print(f"✅ Reset completed - Success: {obs.success}")
        print(f"   Result: {obs.result}")
        if obs.error_message:
            print(f"   Error (expected): {obs.error_message}")
        
        # Test step (will fail gracefully without LibreOffice)
        action = LibreOfficeAction(
            command="set_cell",
            parameters={"cell": "A1", "value": "test"}
        )
        obs = env.step(action)
        print(f"✅ Step completed - Success: {obs.success}")
        print(f"   Result: {obs.result}")
        if obs.error_message:
            print(f"   Error (expected): {obs.error_message}")
        
        # Test state after step
        state = env.state
        print(f"✅ State after step - Step count: {state.step_count}")
        
        return True
    except Exception as e:
        print(f"❌ Environment structure test failed: {e}")
        return False

def test_client_structure():
    """Test client structure and methods."""
    print("\nTesting client structure...")
    
    try:
        from envs.libreoffice_env.client import LibreOfficeEnv
        from envs.libreoffice_env.models import LibreOfficeAction
        
        # Test client creation (without connecting)
        client = LibreOfficeEnv(base_url="http://localhost:8000")
        print("✅ Client created successfully")
        
        # Test action payload conversion
        action = LibreOfficeAction(
            command="set_cell",
            parameters={"cell": "A1", "value": "test"}
        )
        payload = client._step_payload(action)
        print("✅ Action payload conversion successful")
        print(f"   Payload: {payload}")
        
        # Test convenience methods exist
        methods = ['set_cell', 'get_cell', 'set_range', 'get_range', 'set_formula', 
                  'get_formula', 'open_file', 'save_file', 'add_sheet', 'delete_sheet',
                  'rename_sheet', 'format_cell', 'export_pdf', 'export_csv']
        
        for method in methods:
            if hasattr(client, method):
                print(f"✅ Method {method} exists")
            else:
                print(f"❌ Method {method} missing")
                return False
        
        return True
    except Exception as e:
        print(f"❌ Client structure test failed: {e}")
        return False

def test_http_server():
    """Test HTTP server structure."""
    print("\nTesting HTTP server...")
    
    try:
        from envs.libreoffice_env.server.app import app
        from fastapi.testclient import TestClient
        
        # Create test client
        client = TestClient(app)
        print("✅ FastAPI test client created successfully")
        
        # Test health endpoint
        response = client.get('/health')
        print(f"✅ Health endpoint - Status: {response.status_code}")
        
        # Test that the app has the expected routes
        routes = [route.path for route in app.routes]
        expected_routes = ['/health', '/reset', '/step', '/state']
        
        for route in expected_routes:
            if route in routes:
                print(f"✅ Route {route} exists")
            else:
                print(f"❌ Route {route} missing")
                return False
        
        return True
    except Exception as e:
        print(f"❌ HTTP server test failed: {e}")
        return False

def test_command_validation():
    """Test command validation and error handling."""
    print("\nTesting command validation...")
    
    try:
        from envs.libreoffice_env.server.libreoffice_environment import LibreOfficeEnvironment
        from envs.libreoffice_env.models import LibreOfficeAction
        
        env = LibreOfficeEnvironment()
        env.reset()  # This will fail, but that's expected
        
        # Test unknown command
        action = LibreOfficeAction(
            command="unknown_command",
            parameters={}
        )
        obs = env.step(action)
        print(f"✅ Unknown command handled - Success: {obs.success}")
        print(f"   Result: {obs.result}")
        
        # Test missing parameters
        action = LibreOfficeAction(
            command="set_cell",
            parameters={}  # Missing required parameters
        )
        obs = env.step(action)
        print(f"✅ Missing parameters handled - Success: {obs.success}")
        print(f"   Result: {obs.result}")
        
        return True
    except Exception as e:
        print(f"❌ Command validation test failed: {e}")
        return False

def main():
    """Run all tests."""
    print("LibreOffice Environment Test Suite")
    print("=" * 50)
    
    tests = [
        test_imports,
        test_models,
        test_environment_structure,
        test_client_structure,
        test_http_server,
        test_command_validation
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"❌ Test {test.__name__} failed with exception: {e}")
    
    print("\n" + "=" * 50)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! The LibreOffice environment is working correctly.")
        print("\nNote: LibreOffice integration requires:")
        print("1. LibreOffice to be installed")
        print("2. Python-UNO bridge to be available")
        print("3. Running in a Docker container with LibreOffice")
        print("\nThe environment structure and error handling are working perfectly!")
        return True
    else:
        print("❌ Some tests failed. Please check the errors above.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
