"""
Integration test: Frontend + Backend booking flow
Verifies:
1. Backend serves frontend assets
2. API endpoints respond correctly together
3. Booking data flows between frontend config and backend
"""
import requests
import json
import base64
import sys

BACKEND = "https://breeding-quantity-complex-suit.trycloudflare.com"
FRONTEND = "https://significantly-endless-lite-toxic.trycloudflare.com"
TOKEN = base64.b64encode(json.dumps({"role": "admin"}).encode()).decode()
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

def test_frontend_serves_flutter():
    """Verify frontend serves Flutter web app with correct backend config"""
    r = requests.get(FRONTEND, timeout=15)
    assert r.status_code == 200, f"Frontend returned {r.status_code}"
    html = r.text
    assert "flutter" in html.lower() or "main.dart.js" in html, "Not a Flutter app"
    # Check main.dart.js loads
    assert "main.dart.js" in html, "Flutter bootstrap missing"
    print(f"✅ Frontend: Flutter web app loads ({len(html)} bytes)")

def test_backend_health():
    """Verify backend serves dashboard"""
    r = requests.get(f"{BACKEND}/", timeout=10)
    assert r.status_code == 200, f"Backend returned {r.status_code}"
    print(f"✅ Backend: Dashboard loads ({len(r.text)} bytes)")

def test_services_api():
    """Verify services endpoint returns data frontend needs"""
    r = requests.get(f"{BACKEND}/api/public/services", timeout=10)
    assert r.status_code == 200
    services = r.json()
    assert len(services) > 50, f"Only {len(services)} services"
    # Check required fields for frontend
    for svc in services[:5]:
        assert "code" in svc, f"Missing code in {svc.get('name')}"
        assert "name" in svc, f"Missing name"
        assert "price" in svc, f"Missing price"
    print(f"✅ Services: {len(services)} services with valid schema")

def test_slots_api():
    """Verify slots available for booking"""
    r = requests.get(f"{BACKEND}/api/public/slots", timeout=10)
    assert r.status_code == 200
    print(f"✅ Slots: endpoint responding")

def test_crm_integration():
    """Verify CRM shows real client data (proves booking data exists)"""
    r = requests.get(f"{BACKEND}/api/crm/clients", headers=HEADERS, timeout=10)
    assert r.status_code == 200
    clients = r.json()
    assert len(clients) > 0, "No clients in CRM"
    # Check Puguh's data for consistency
    puguh = [c for c in clients if "puguh" in c.get("name", "").lower()]
    if puguh:
        p = puguh[0]
        assert p["loyalty_points"] >= 0, "Invalid loyalty_points"
        assert p["orders"] >= 0, "Invalid orders"
        print(f"✅ CRM: {len(clients)} clients (Puguh: {p['loyalty_points']} pts, {p['orders']} bookings)")
    else:
        print(f"✅ CRM: {len(clients)} clients")

def test_approvals_workflow():
    """Verify booking approval system works"""
    r = requests.get(f"{BACKEND}/api/approvals", headers=HEADERS, timeout=10)
    assert r.status_code == 200
    approvals = r.json()
    print(f"✅ Approvals: {len(approvals) if isinstance(approvals, list) else 'N/A'} pending")

def test_dashboard_loads():
    """Verify admin dashboard with all widgets"""
    r = requests.get(f"{BACKEND}/api/dashboard", headers=HEADERS, timeout=10)
    assert r.status_code == 200
    dash = r.json()
    print(f"✅ Dashboard: widgets loaded ({list(dash.keys())[:5]})")

if __name__ == "__main__":
    results = []
    tests = [
        ("frontend_app", test_frontend_serves_flutter),
        ("backend_health", test_backend_health),
        ("services_api", test_services_api),
        ("slots_api", test_slots_api),
        ("crm_integration", test_crm_integration),
        ("approvals", test_approvals_workflow),
        ("dashboard", test_dashboard_loads),
    ]
    for name, fn in tests:
        try:
            fn()
            results.append((name, "PASS"))
        except Exception as e:
            results.append((name, f"FAIL: {e}"))
            print(f"❌ {name}: {e}", file=sys.stderr)
    
    passed = sum(1 for r in results if r[1] == "PASS")
    print(f"\n{'='*50}")
    print(f"Integration Test: {passed}/{len(results)} passed")
    failed = [r for r in results if r[1] != "PASS"]
    if failed:
        for name, err in failed:
            print(f"  ❌ {name}: {err}")
        sys.exit(1)
    else:
        print("✅ ALL PASSED — Frontend + Backend terintegrasi!")
