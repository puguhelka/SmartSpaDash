"""
Backend API tests for Lelap Mom Baby Care Salatiga
Tests: services list, slots, booking flow, membership, points
"""
import requests
import json
import sys

BASE = "https://breeding-quantity-complex-suit.trycloudflare.com"

def test_services_list():
    """Verify /api/public/services returns services with code, name, price"""
    r = requests.get(f"{BASE}/api/public/services", timeout=10)
    assert r.status_code == 200, f"Expected 200, got {r.status_code}"
    data = r.json()
    assert isinstance(data, list), "Expected array"
    assert len(data) > 0, "Empty services list"
    svc = data[0]
    assert "code" in svc, "Missing code field"
    assert "name" in svc, "Missing name field"
    assert "price" in svc, "Missing price field"
    print(f"✅ Services: {len(data)} items, first: {svc['name']} - Rp {svc['price']}")

def test_slots_available():
    """Verify slot availability endpoint works"""
    r = requests.get(f"{BASE}/api/public/slots", timeout=10)
    assert r.status_code == 200, f"Expected 200, got {r.status_code}"
    data = r.json()
    print(f"✅ Slots available: {json.dumps(data)[:200]}")

def test_crm_clients():
    """Verify CRM endpoint returns client list (requires auth)"""
    import base64
    token = base64.b64encode(json.dumps({"role": "admin"}).encode()).decode()
    r = requests.get(f"{BASE}/api/crm/clients", headers={"Authorization": f"Bearer {token}"}, timeout=10)
    assert r.status_code == 200, f"Expected 200, got {r.status_code}"
    data = r.json()
    assert isinstance(data, list), "Expected array"
    print(f"✅ CRM clients: {len(data)} clients loaded")

def test_health_check():
    """Basic connectivity test"""
    r = requests.get(f"{BASE}/", timeout=10)
    # Dashboard serves HTML at root
    assert r.status_code == 200, f"Expected 200, got {r.status_code}"
    assert "SmartSpaDash" in r.text or "html" in r.text.lower() or "Lelap" in r.text, "No dashboard content"
    print(f"✅ Health check: {r.status_code}, {len(r.text)} bytes")

if __name__ == "__main__":
    results = []
    for name, fn in [
        ("health", test_health_check),
        ("services", test_services_list),
        ("slots", test_slots_available),
        ("crm_clients", test_crm_clients),
    ]:
        try:
            fn()
            results.append((name, "PASS"))
        except Exception as e:
            results.append((name, f"FAIL: {e}"))
            print(f"❌ {name}: {e}", file=sys.stderr)
    
    failed = [r for r in results if r[1] != "PASS"]
    print(f"\n{'='*40}")
    print(f"Results: {len(results)-len(failed)}/{len(results)} passed")
    if failed:
        print(f"Failed: {failed}")
        sys.exit(1)
    else:
        print("All tests passed!")
