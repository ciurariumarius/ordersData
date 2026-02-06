# Orders Data Export Automation

This Google Apps Script project automates the export of order data from Shopify and Gomag into a Google Sheet. It fetches detailed order information including transaction IDs, values, shipping costs, and product details.

## Features
- **Shopify Export**: Fetches orders from Shopify Admin API.
- **Gomag Export**: Fetches orders from Gomag API.
- **WooCommerce Export**: Fetches orders from WooCommerce REST API.
- **Dynamic Configuration**: Reads API credentials and settings from a "Config" sheet.
- **Robustness**: Handles API rate limits and pagination automatically.
- **Clean Output**: Formats data into clear, readable columns.

## Prerequisites
- A Google Sheet to host the script.
- API Credentials for Shopify (Admin API Access Token).
- API Credentials for Gomag (API Key).
- API Credentials for WooCommerce (Consumer Key & Consumer Secret).

## Setup

### 1. Create the Config Sheet
Create a new tab in your Google Sheet named exactly `Config`.
Add the following key-value pairs in **Column A** and **Column B**:

| Column A (Key) | Column B (Value) | Description |
| :--- | :--- | :--- |
| `shopifyDomain` | `your-shop.myshopify.com` | Your Shopify Admin Domain |
| `shopifyAPIkey` | `shpat_xxxxxxxxxxxx` | Your Shopify Access Token |
| `gomagDomain` | `https://www.your-site.ro` | Your Gomag Shop URL (no trailing slash) |
| `gomagAPIkey` | `xxxxxxxxxxxxxxxxx` | Your Gomag API Key |
| `wooUrl` | `https://www.your-woo-store.com` | Your WooCommerce Store URL |
| `wooConsumerKey` | `ck_xxxxxxxxxxxx` | WooCommerce Consumer Key |
| `wooConsumerSecret` | `cs_xxxxxxxxxxxx` | WooCommerce Consumer Secret |
| `Days` | `30` | Number of days to look back for orders (Default: 30) |

### 2. Script Files
Ensure the following files are present in your Apps Script project:
- `Config.gs`: Handles reading the configuration.
- `Utilities.gs`: Shared helper functions for API requests and Sheet operations.
- `ShopifyOrders.gs`: Logic for Shopify exports.
- `GomagOrders.gs`: Logic for Gomag exports.
- `WooCommerceOrders.gs`: Logic for WooCommerce exports.

## How to Run

### Shopify Export
1. Select the function `runShopifyOrderExport`.
2. Click **Run**.
3. Check the **ShopifyOrders** tab.

### Gomag Export
1. Select the function `runGomagOrderExport`.
2. Click **Run**.
3. Check the **GomagOrders** tab.

### WooCommerce Export
1. Select the function `runWooCommerceOrderExport`.
2. Click **Run**.
3. Check the **WooCommerceOrders** tab.

## Output Columns
The scripts generate the following columns:
1.  **Date**
2.  **Transaction ID**
3.  **Value** (Total Order Value)
4.  **Shipping** (Transport Cost)
5.  **Items Revenue** (Sum of Product Prices * Quantity)
6.  **Product IDs** (Comma-separated list)
7.  **Status**

## Troubleshooting
- **Sheet "Config" not found**: Ensure the tab is named exactly `Config` (case-sensitive).
- **Missing credentials**: Check the script logs for errors about missing keys in the Config sheet.
- **Rate Limit Exceeded**: The scripts have built-in retry logic, but if you process a huge volume, you might need to run it in smaller batches (reduce "Days").
