# BrowserStack Testathon - Comprehensive Test Plan & Test Cases

This document details the test scenarios, folder taxonomy, and test cases designed for the **BrowserStack Testathon**, covering all critical user flows, edge cases, and cross-browser requirements.

---

## 1. Test Suite Architecture & Folder Structure

```text
BrowserStack Test Management Project: "BrowserStack Testathon"
├── [Folder 1] Authentication & Session Management
├── [Folder 2] Product Catalog & Vendor Filtering
├── [Folder 3] Cart Management & Counter Verification
├── [Folder 4] Checkout & Order Fulfillment
├── [Folder 5] Negative Scenarios & Edge Cases
└── [Folder 6] Cloud Infrastructure & Local Tunneling
```

---

## 2. Documented Test Cases

### Folder 1: Authentication & Session Management

#### `TC-AUTH-001`: Valid User Authentication (Happy Path)
* **Priority**: P0 (Blocker)
* **Pre-conditions**: User is on `https://bstackdemo.com/signin`.
* **Test Steps**:
  1. Click on the "Select Username" dropdown.
  2. Choose `demouser`.
  3. Click on the "Select Password" dropdown.
  4. Choose `testingisfun99`.
  5. Click the "Log In" button (`#login-btn`).
* **Expected Result**: User is successfully authenticated; the navbar displays the username `demouser`, and user is redirected to product catalog.
* **Automation Mapping**: `bstackdemo_suite.py` -> `test_01_user_login_positive()`

#### `TC-AUTH-002`: User Logout & Session Invalidation
* **Priority**: P1 (Critical)
* **Pre-conditions**: User is logged in as `demouser`.
* **Test Steps**:
  1. Locate and click the "Logout" link in the top right navbar.
* **Expected Result**: Session is terminated; navbar reverts to showing "Sign In"; protected checkout redirects to `/signin`.

---

### Folder 2: Product Catalog & Vendor Filtering

#### `TC-CAT-001`: Vendor Filtering by Manufacturer (Apple)
* **Priority**: P1 (Critical)
* **Pre-conditions**: User is on `https://bstackdemo.com`.
* **Test Steps**:
  1. In the left sidebar filter list, select the "Apple" checkbox.
  2. Wait for the product shelf animation to complete.
* **Expected Result**: The product catalog dynamically re-renders; all displayed products belong to Apple; product count matches expected Apple inventory.
* **Automation Mapping**: `bstackdemo_suite.py` -> `test_02_product_filtering()`

#### `TC-CAT-002`: Multi-Vendor Filtering (Apple + Samsung)
* **Priority**: P2 (Major)
* **Test Steps**:
  1. Select both "Apple" and "Samsung" filter checkboxes.
* **Expected Result**: Catalog displays union of Apple and Samsung products; other brands remain hidden.

---

### Folder 3: Cart Management & Counter Verification

#### `TC-CART-001`: Add Single Item to Shopping Cart
* **Priority**: P0 (Blocker)
* **Pre-conditions**: User is on product catalog.
* **Test Steps**:
  1. Click "Add to cart" on the first listed product (e.g., iPhone 12).
* **Expected Result**:
  1. Floating cart drawer (`.float-cart--open`) slides into view.
  2. Cart badge counter increments to `1`.
  3. Product title, unit price, and subtotal reflect the added item.
* **Automation Mapping**: `bstackdemo_suite.py` -> `test_03_add_to_cart()`

#### `TC-CART-002`: Item Removal from Cart
* **Priority**: P1 (Critical)
* **Test Steps**:
  1. In the open cart drawer, click the "X" (delete) icon next to the item.
* **Expected Result**: Item is removed; subtotal recalculates to `$0.00`; cart badge displays `0`.

---

### Folder 4: Checkout & Order Fulfillment

#### `TC-CHK-001`: Complete End-to-End Checkout Flow
* **Priority**: P0 (Blocker)
* **Pre-conditions**: User is logged in with at least 1 item in the cart.
* **Test Steps**:
  1. Click the "Checkout" button in the cart drawer.
  2. Fill in the shipping form:
     * First Name: `Alex`
     * Last Name: `Smith`
     * Address: `100 Innovation Parkway`
     * Province/State: `California`
     * Postal Code: `94016`
  3. Click "Submit" (`#checkout-shipping-continue`).
* **Expected Result**: Order is submitted successfully; confirmation message displays: `"Your Order has been successfully placed."`; order receipt is generated.
* **Automation Mapping**: `bstackdemo_suite.py` -> `test_04_checkout_order_placement()`

---

### Folder 5: Negative Scenarios & Edge Cases

#### `TC-NEG-001`: Authentication with Invalid Password
* **Priority**: P1 (Critical)
* **Pre-conditions**: User is on `https://bstackdemo.com/signin`.
* **Test Steps**:
  1. Select valid username `demouser`.
  2. Select incorrect/mismatched password option or trigger invalid credentials.
  3. Click "Log In".
* **Expected Result**: Login fails; error toast `.api-error` is displayed with message `"Invalid Password"`; user remains on signin page.
* **Automation Mapping**: `bstackdemo_suite.py` -> `test_05_negative_login_invalid_password()`

#### `TC-NEG-002`: Checkout Attempt with Empty Cart
* **Priority**: P2 (Major)
* **Pre-conditions**: Cart contains 0 items.
* **Test Steps**:
  1. Open empty cart drawer.
  2. Attempt to click checkout button.
* **Expected Result**: Checkout button is disabled or alerts user: `"Add some products in the bag!"`.

---

### Folder 6: Cloud Infrastructure & Local Tunneling

#### `TC-TUN-001`: BrowserStack Local Tunnel Routing (Intranet API & Portal)
* **Priority**: P1 (Critical)
* **Pre-conditions**: `browserstack-local` tunnel active on test host.
* **Test Steps**:
  1. Direct remote cloud browser on Windows 11 Chrome to `http://bs-local.com:8888`.
  2. Measure tunnel latency.
  3. Verify local server HTML response and status badge.
* **Expected Result**: Remote cloud browser accesses localhost portal cleanly through the encrypted tunnel; tunnel latency < 50ms; status code 200.
* **Automation Mapping**: `test_runner.py` -> `UI Web Navigation - Local Tunnel Intranet`
