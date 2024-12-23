clientId = "1000.I4B8E3BTIR0ZJA7HKO4HJ0HXQ62XZG"
clientSecret = "33aeec3188483ca6b4154ad677ac9e8f123aa25fca"
grantType = "refresh_token"
inventory_refresh_token = (
    "1000.46ff94d18b6a7158cd7420329e577e8f.f7a056f71562e01956d6dede5204427f"
)
books_refresh_token = (
    "1000.55c5cdc6a3db5af5d0e5f09a01c480ec.941ab8dc22daa5bc6aa3e3c74bd477b8"
)
org_id = "776755316"

#Authentication
INVENTORY_URL = "https://accounts.zoho.com/oauth/v2/token?refresh_token={inventory_refresh_token}&client_id={clientId}&client_secret={clientSecret}&redirect_uri=http://www.zoho.com/inventory&grant_type={grantType}"

BOOKS_URL = "https://accounts.zoho.com/oauth/v2/token?refresh_token={books_refresh_token}&client_id={clientId}&client_secret={clientSecret}&redirect_uri=http://www.zoho.com/inventory&grant_type={grantType}"

#URLS FOR DATA FETCHING
PURCHASE_URL = "https://books.zoho.com/api/v3/purchaseorders?search_text={search_text}&page={page}&per_page=200&sort_order=D&sort_column=date&organization_id={org_id}"

PURCHASE_ORDER_URL = "https://books.zoho.com/api/v3/purchaseorders/{purchase_order_id}?organization_id={org_id}"

ITEM_URL = "https://www.zohoapis.com/books/v3/items?search_text={search_text}&page=1&per_page=2000&organization_id={org_id}"