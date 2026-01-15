
import sys
import os
# Add current directory to path
sys.path.append(os.getcwd())

from manning_web_app import get_category

def test_fuzzy():
    print("Testing Fuzzy Mapping Logic...\n")
    
    cases = [
        # Location, Role, Expected Station
        ("ikes", "random cold prep string", "SALAD BAR"),
        ("ikes", "weird deli text", "HOMESLICE"),
        ("ikes", "some flips cook", "GRILL"), 
        ("ikes", "my production cook", "HOMESTYLE"),
        ("southside", "random pizza place", "POTOMAC PIE"),
        ("southside", "some pasta dish", "LITTLE ITALY"),
        ("southside", "cleaning utility", "DISHROOM"),
        ("southside", "new dessert item", "BLUE RIDGE BAKERY"),
    ]
    
    passed = 0
    for loc, role, expected in cases:
        result = get_category(role, loc)
        status = "PASS" if result == expected else f"FAIL (Got {result})"
        print(f"[{loc.upper()}] '{role}' -> '{expected}' : {status}")
        if result == expected:
            passed += 1
            
    print(f"\n{passed}/{len(cases)} tests passed.")

if __name__ == "__main__":
    test_fuzzy()
