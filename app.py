from flask import Flask, render_template, send_file
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, timezone
import threading
import time

# UTC+2 timezone
UTC_PLUS_2 = timezone(timedelta(hours=2))

app = Flask(__name__)

# Make format_onhand_display available to templates
@app.context_processor
def utility_processor():
    return dict(format_onhand=format_onhand_display)

# Cache storage
cache = {
    'data': None,
    'timestamp': None,
    'lock': threading.Lock()
}
CACHE_DURATION = timedelta(minutes=5)  # Cache expires after 5 minutes

# API Configuration from Postman collection
API_URL = "http://192.168.5.90:8090/IQRetailRestAPI/v1/IQ_API_Request_Stock_Attributes?callformat=xml"
API_HEADERS = {
    "Authorization": "Basic OTk6QVBJVGVzdA==",
    "Content-Type": "application/json"
}
API_BODY = """<IQ_API>
	<IQ_API_Request_Stock>
		<IQ_Company_Number>F01</IQ_Company_Number>
		<IQ_Terminal_Number>22</IQ_Terminal_Number>
		<IQ_User_Number>52</IQ_User_Number>
		<IQ_User_Password>1E68EEE44919E675F5A8A455A755789099B54E4D</IQ_User_Password>
		<IQ_Partner_Passphrase>743B25C6C57BA9A4D02EBBAD9D11B9ADC47A1BCA</IQ_Partner_Passphrase>
	</IQ_API_Request_Stock>
</IQ_API>"""


def find_parent_elements(root):
    """Build a parent mapping for all elements"""
    parent_map = {c: p for p in root.iter() for c in p}
    return parent_map


def fetch_stock_data():
    """Fetch stock data from the API and parse XML response"""
    try:
        response = requests.post(API_URL, headers=API_HEADERS, data=API_BODY, timeout=30)
        response.raise_for_status()
        
        # Parse XML response
        root = ET.fromstring(response.content)
        
        products = []
        
        # Build parent mapping
        parent_map = find_parent_elements(root)
        
        # Strategy: Find all elements with our target field names and group by their parent
        parent_products = {}
        
        for elem in root.iter():
            # Remove namespace prefixes if present
            tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            
            if tag_name in ['Description', 'Supplier_Item_Code', 'Onhand_Available']:
                # Find the parent element that likely represents a product/item
                current = elem
                item_element = None
                
                # Look for a parent element that might represent an item/product
                # Check up to 3 levels up
                for level in range(3):
                    if current in parent_map:
                        parent = parent_map[current]
                        parent_tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                        
                        # Check if this looks like an item/product container
                        if any(keyword in parent_tag.lower() for keyword in ['item', 'product', 'stock', 'record', 'row', 'line']):
                            item_element = parent
                            break
                        current = parent
                    else:
                        break
                
                # Use the found item element, or use the direct parent if no item element found
                key_elem = item_element if item_element else (parent_map.get(elem, elem))
                key_id = id(key_elem)
                
                if key_id not in parent_products:
                    parent_products[key_id] = {
                        'Description': '',
                        'Supplier_Item_Code': '',
                        'Onhand_Available': ''
                    }
                
                parent_products[key_id][tag_name] = elem.text or '' if elem.text else ''
        
        # Convert to list
        if parent_products:
            products = list(parent_products.values())
        else:
            # Fallback: collect all fields and match by position
            descriptions = []
            supplier_codes = []
            onhand_values = []
            
            for elem in root.iter():
                tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
                
                if tag_name == 'Description':
                    descriptions.append(elem.text or '')
                elif tag_name == 'Supplier_Item_Code':
                    supplier_codes.append(elem.text or '')
                elif tag_name == 'Onhand_Available':
                    onhand_values.append(elem.text or '')
            
            # Match by position
            max_len = max(len(descriptions), len(supplier_codes), len(onhand_values))
            for i in range(max_len):
                desc = descriptions[i] if i < len(descriptions) else ''
                code = supplier_codes[i] if i < len(supplier_codes) else ''
                onhand = onhand_values[i] if i < len(onhand_values) else ''
                
                if desc or code or onhand:
                    products.append({
                        'Description': desc,
                        'Supplier_Item_Code': code,
                        'Onhand_Available': onhand
                    })
        
        # Remove empty products
        products = [p for p in products if any(str(v).strip() for v in p.values())]
        
        return products
    
    except requests.exceptions.RequestException as e:
        return {'error': f'API Request Error: {str(e)}'}
    except ET.ParseError as e:
        error_msg = f'XML Parse Error: {str(e)}'
        if 'response' in locals():
            try:
                error_msg += f'\nResponse preview: {response.content[:500].decode("utf-8", errors="ignore")}'
            except:
                pass
        return {'error': error_msg}
    except Exception as e:
        return {'error': f'Unexpected Error: {str(e)}'}


def get_water_feature_components():
    """Returns a dictionary mapping water feature codes to their components.
    Each value is a list of tuples: (component_code, quantity)"""
    return {
        'WJVBU55E': [('JVBU55E', 1), ('JPRD35E', 1)],
        'WJVBU55S': [('JVBU55S', 1), ('JPRD35S', 1)],
        'WJVBI75PA': [('JVBI75PA', 1), ('JPRD57PA', 1)],
        'WJVBI75C': [('JVBI75C', 1), ('JPRD57C', 1)],
        'WJVBI75AC': [('JVBI75AC', 1), ('JPRD57AC', 1)],
        'WJVBI75E': [('JVBI75E', 1), ('JPRD57E', 1)],
        'WJVBI75B': [('JVBI75B', 1), ('JPRD57B', 1)],
        'WJVBI75S': [('JVBI75S', 1), ('JPRD57S', 1)],
        'WJVBI75O': [('JVBI75O', 1), ('JPRD57O', 1)],
        'WJVBI55PA': [('JVBI55PA', 1), ('JPRD41PA', 1)],
        'WJVBI55C': [('JVBI55C', 1), ('JPRD41C', 1)],
        'WJVBI55AC': [('JVBI55AC', 1), ('JPRD41AC', 1)],
        'WJVBI55E': [('JVBI55E', 1), ('JPRD41E', 1)],
        'WJVBI55B': [('JVBI55B', 1), ('JPRD41B', 1)],
        'WJVBI55S': [('JVBI55S', 1), ('JPRD41S', 1)],
        'WJVBI55O': [('JVBI55O', 1), ('JPRD41O', 1)],
        'WJVBI43PA': [('JVBI43PA', 1), ('JPRD25PA', 1)],
        'WJVBI43C': [('JVBI43C', 1), ('JPRD25C', 1)],
        'WJVBI43AC': [('JVBI43AC', 1), ('JPRD25AC', 1)],
        'WJVBI43E': [('JVBI43E', 1), ('JPRD25E', 1)],
        'WJVBI43B': [('JVBI43B', 1), ('JPRD25B', 1)],
        'WJVBI43S': [('JVBI43S', 1), ('JPRD25S', 1)],
        'WJVBI42PA': [('JVBI42PA', 1), ('JPRD25PA', 1)],
        'WJVBI42C': [('JVBI42C', 1), ('JPRD25C', 1)],
        'WJVBI42AC': [('JVBI42AC', 1), ('JPRD25AC', 1)],
        'WJVBI42E': [('JVBI42E', 1), ('JPRD25E', 1)],
        'WJVBI42B': [('JVBI42B', 1), ('JPRD25B', 1)],
        'WJVBI42S': [('JVBI42S', 1), ('JPRD25S', 1)],
        'WJVBI42O': [('JVBI42O', 1), ('JPRD25O', 1)],
        'WJVEO255E': [('JVEO55E', 2)],
        'WJVEO255O': [('JVEO55O', 2)],
        'WJVEO255B': [('JVEO55B', 2)],
        'WJVEO355E': [('JVEO55E', 3)],
        'WJVEO355O': [('JVEO55O', 3)],
        'WJVEO355B': [('JVEO55B', 3)],
        'WJVEO375E': [('JVEO55E', 2), ('JVEO75E', 1)],
        'WJVEO375O': [('JVEO55O', 2), ('JVEO75E', 1)],
        'WJVEO375B': [('JVEO55B', 2), ('JVEO75E', 1)],
        'WJVEO55E': [('JVEO55E', 1), ('JPRD46E', 1)],
        'WJVEO55B': [('JVEO55B', 1), ('JPRD46B', 1)],
        'WJVEO55O': [('JVEO55O', 1), ('JPRD46O', 1)],
        'WJVEV33C': [('JVEV33C', 1), ('JPRD25C', 1)],
        'WJVEV33E': [('JVEV33E', 1), ('JPRD25E', 1)],
        'WJVEV33B': [('JVEV33B', 1), ('JPRD25B', 1)],
        'WJVEV33S': [('JVEV33S', 1), ('JPRD25S', 1)],
        'WJVEV33O': [('JVEV33O', 1), ('JPRD25O', 1)],
        'WJVLC50C': [('JVLC50C', 1), ('JPRD41C', 1)],
        'WJVLC50E': [('JVLC50E', 1), ('JPRD41E', 1)],
        'WJVOP41E': [('JVOP41E', 1), ('JPRD30E', 1)],
        'WJVOP41B': [('JVOP41B', 1), ('JPRD30B', 1)],
        'WJVRL45E': [('JVRL45E', 1), ('JPRD41E', 1)],
        'WJVRL45B': [('JVRL45B', 1), ('JPRD41B', 1)],
        'WJVRR57PA': [('JVRR57PA', 1), ('JPRD46PA', 1)],
        'WJVRR57C': [('JVRR57C', 1), ('JPRD46C', 1)],
        'WJVRR57E': [('JVRR57E', 1), ('JPRD46E', 1)],
        'WJVRR57B': [('JVRR57B', 1), ('JPRD46B', 1)],
        'WJVRR57S': [('JVRR57S', 1), ('JPRD46S', 1)],
        'WJVRR57O': [('JVRR57O', 1), ('JPRD46O', 1)],
        'WJVRR33PA': [('JVRR33PA', 1), ('JPRD25PA', 1)],
        'WJVRR33C': [('JVRR33C', 1), ('JPRD25C', 1)],
        'WJVRR33E': [('JVRR33E', 1), ('JPRD25E', 1)],
        'WJVRR33B': [('JVRR33B', 1), ('JPRD25B', 1)],
        'WJVRR33S': [('JVRR33S', 1), ('JPRD25S', 1)],
        'WJVRR33O': [('JVRR33O', 1), ('JPRD25O', 1)],
    }


def calculate_water_feature_onhand(water_feature_code, all_products_dict):
    """Calculate on-hand available quantity for a water feature based on its components.
    
    Args:
        water_feature_code: The supplier code of the water feature (e.g., 'WJVBU55E')
        all_products_dict: Dictionary mapping supplier codes to their on-hand quantities
    
    Returns:
        Integer representing the calculated on-hand quantity
    """
    components = get_water_feature_components().get(water_feature_code.upper())
    if not components:
        return 0
    
    # Group components by code to handle multiple quantities of the same component
    component_quantities = {}
    for component_code, quantity in components:
        if component_code not in component_quantities:
            component_quantities[component_code] = 0
        component_quantities[component_code] += quantity
    
    # Calculate available quantities for each component type
    available_quantities = []
    for component_code, required_qty in component_quantities.items():
        # Get on-hand quantity for this component
        onhand_str = all_products_dict.get(component_code, '0')
        try:
            onhand = int(float(str(onhand_str).strip() or '0'))
        except (ValueError, TypeError):
            onhand = 0
        
        # If component doesn't exist, return 0
        if component_code not in all_products_dict:
            return 0
        
        # Divide by required quantity and round down (no decimals)
        available = onhand // required_qty
        available_quantities.append(available)
    
    # Return the minimum (bottleneck component)
    return min(available_quantities) if available_quantities else 0


def get_cached_stock_data():
    """Get stock data from cache or fetch new data if cache is stale/missing.
    Returns tuple: (data, timestamp, is_cached, api_error) where:
    - is_cached indicates if cached data was used
    - api_error indicates if we're using cached data due to an API error"""
    with cache['lock']:
        now = datetime.now(UTC_PLUS_2)
        
        # Check if cache is valid
        if cache['data'] is not None and cache['timestamp'] is not None:
            if now - cache['timestamp'] < CACHE_DURATION:
                # Cache is still valid, return it (no API error)
                return cache['data'], cache['timestamp'], True, False
        
        # Cache is stale or missing, try to fetch new data
        result = fetch_stock_data()
        
        # Check if API call failed (result is a dict with 'error' key)
        if isinstance(result, dict) and 'error' in result:
            # API call failed, return cached data if available
            if cache['data'] is not None and cache['timestamp'] is not None:
                return cache['data'], cache['timestamp'], True, True
            else:
                # No cached data available, return error
                return result, None, False, True
        
        # API call succeeded, update cache
        cache['data'] = result
        cache['timestamp'] = now
        
        return result, now, False, False


def update_cache_periodically():
    """Background thread function to update cache every 5 minutes"""
    while True:
        try:
            time.sleep(CACHE_DURATION.total_seconds())
            # Fetch new data and update cache only if successful
            with cache['lock']:
                result = fetch_stock_data()
                # Only update cache if API call succeeded (not an error)
                if not (isinstance(result, dict) and 'error' in result):
                    cache['data'] = result
                    cache['timestamp'] = datetime.now(UTC_PLUS_2)
                    print(f"Cache updated at {cache['timestamp']}")
                else:
                    print(f"Cache update failed: {result.get('error', 'Unknown error')}. Keeping existing cache.")
        except Exception as e:
            print(f"Error updating cache: {str(e)}. Keeping existing cache.")


def format_onhand_display(onhand_value, is_staff=False):
    """Format on-hand available quantity for display.
    
    Args:
        onhand_value: The on-hand quantity (string or number)
        is_staff: If True, show actual values including negatives. If False, show "Out Of Stock" for 0 or less.
    
    Returns:
        Formatted string for display
    """
    if is_staff:
        # Staff view: show actual values including negatives
        return str(onhand_value) if onhand_value else '-'
    else:
        # Normal view: show "Out Of Stock" for 0 or less (with styling class)
        try:
            onhand_num = int(float(str(onhand_value).strip() or '0'))
            if onhand_num <= 0:
                return '<span class="out-of-stock">Out Of Stock</span>'
            return str(onhand_num)
        except (ValueError, TypeError):
            return '<span class="out-of-stock">Out Of Stock</span>'


@app.route('/')
def index():
    """Main page - fetch and display stock data"""
    result, cache_timestamp, is_cached, api_error = get_cached_stock_data()
    
    # If there's an error and no cached data, return it
    if isinstance(result, dict) and 'error' in result:
        return render_template('index.html', error=result['error'], japi_products=[], dropin_products=[], driptrays_products=[], water_features_products=[], nb_fx_products=[], all_products=[], cache_timestamp=None, is_cached=False, api_error=False)
    
    # Convert to list if needed
    if not isinstance(result, list):
        result = []
    
    # Separate products into categories:
    # 1. Water Features (Supplier_Item_Code starts with "W")
    # 2. Organic Feature Pots (Supplier_Item_Code starts with "NBP", "FX", "NR", or "RV", or Description contains "Organic Fiberglass")
    # 3. Driptrays (Supplier_Item_Code contains "JPRD" or "JPQD")
    # 4. Drop-in Pots (Description contains "Styler" or "Shrub")
    # 5. Japi Planters (Supplier_Item_Code starts with "J" or "NSGB2")
    # 6. Plastic Woven Mats (Supplier_Item_Code starts with "PWM")
    # 7. All Products (all products from API - for separate tab)
    water_features_products = []
    nb_fx_products = []
    driptrays_products = []
    dropin_products = []
    japi_products = []
    plasticmats_products = []
    
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        # Skip products without a supplier item code
        if not supplier_code:
            continue
        description = (product.get('Description') or '').strip().upper()
        
        # Priority 1: Water Features (Supplier_Item_Code starts with "W")
        if supplier_code.startswith('W'):
            water_features_products.append(product)
        # Priority 2: Organic Feature Pots (Supplier_Item_Code starts with "NBP", "FX", "NR", or "RV", or Description contains "Organic Fiberglass")
        elif supplier_code.startswith('NBP') or supplier_code.startswith('FX') or supplier_code.startswith('NR') or supplier_code.startswith('RV') or 'ORGANIC FIBERGLASS' in description:
            nb_fx_products.append(product)
        # Priority 3: Driptrays (Supplier_Item_Code contains "JPRD" or "JPQD")
        elif 'JPRD' in supplier_code or 'JPQD' in supplier_code:
            driptrays_products.append(product)
        # Priority 4: Drop-in Pots (Description contains "Styler" or "Shrub")
        elif 'STYLER' in description or 'SHRUB' in description:
            dropin_products.append(product)
        # Priority 5: Japi Planters (Supplier_Item_Code starts with "J" or "NSGB2")
        elif supplier_code.startswith('J') or supplier_code.startswith('NSGB2'):
            japi_products.append(product)
        # Priority 6: Plastic Woven Mats (Supplier_Item_Code starts with "PWM")
        elif supplier_code.startswith('PWM'):
            plasticmats_products.append(product)
    
    # Create a dictionary mapping supplier codes to on-hand quantities for water feature calculations
    products_dict = {}
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        if supplier_code:
            products_dict[supplier_code] = product.get('Onhand_Available', '0')
    
    # Calculate on-hand quantities for water features
    for product in water_features_products:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
        product['Onhand_Available'] = str(calculated_onhand)
    
    # Sort all lists alphabetically by Description
    japi_products = sorted(japi_products, key=lambda x: (x.get('Description') or '').lower())
    dropin_products = sorted(dropin_products, key=lambda x: (x.get('Description') or '').lower())
    driptrays_products = sorted(driptrays_products, key=lambda x: (x.get('Description') or '').lower())
    water_features_products = sorted(water_features_products, key=lambda x: (x.get('Description') or '').lower())
    nb_fx_products = sorted(nb_fx_products, key=lambda x: (x.get('Description') or '').lower())
    plasticmats_products = sorted(plasticmats_products, key=lambda x: (x.get('Description') or '').lower())
    
    # Filter all_products to exclude products with certain words in description
    excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
    all_products_filtered = []
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip()
        # Skip products without a supplier item code
        if not supplier_code:
            continue
        description = (product.get('Description') or '').strip()
        # Check if description contains any excluded words (case-insensitive)
        if not any(word.lower() in description.lower() for word in excluded_words):
            all_products_filtered.append(product)
    
    # Calculate on-hand quantities for water features in all_products
    for product in all_products_filtered:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        if supplier_code.startswith('W'):
            calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
            product['Onhand_Available'] = str(calculated_onhand)
    
    all_products = sorted(all_products_filtered, key=lambda x: (x.get('Description') or '').lower())
    
    return render_template('index.html', error=None, japi_products=japi_products, dropin_products=dropin_products, driptrays_products=driptrays_products, water_features_products=water_features_products, nb_fx_products=nb_fx_products, plasticmats_products=plasticmats_products, all_products=all_products, cache_timestamp=cache_timestamp, is_cached=is_cached, api_error=api_error, is_staff=False)


@app.route('/staff')
def staff():
    """Staff page - same as main page but shows actual values including negatives"""
    result, cache_timestamp, is_cached, api_error = get_cached_stock_data()
    
    # If there's an error and no cached data, return it
    if isinstance(result, dict) and 'error' in result:
        return render_template('index.html', error=result['error'], japi_products=[], dropin_products=[], driptrays_products=[], water_features_products=[], nb_fx_products=[], all_products=[], cache_timestamp=None, is_cached=False, api_error=False, is_staff=True)
    
    # Convert to list if needed
    if not isinstance(result, list):
        result = []
    
    # Separate products into categories:
    # 1. Water Features (Supplier_Item_Code starts with "W")
    # 2. Organic Feature Pots (Supplier_Item_Code starts with "NBP", "FX", "NR", or "RV", or Description contains "Organic Fiberglass")
    # 3. Driptrays (Supplier_Item_Code contains "JPRD" or "JPQD")
    # 4. Drop-in Pots (Description contains "Styler" or "Shrub")
    # 5. Japi Planters (Supplier_Item_Code starts with "J" or "NSGB2")
    # 6. Plastic Woven Mats (Supplier_Item_Code starts with "PWM")
    # 7. All Products (all products from API - for separate tab)
    water_features_products = []
    nb_fx_products = []
    driptrays_products = []
    dropin_products = []
    japi_products = []
    plasticmats_products = []
    
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        # Skip products without a supplier item code
        if not supplier_code:
            continue
        description = (product.get('Description') or '').strip().upper()
        
        # Priority 1: Water Features (Supplier_Item_Code starts with "W")
        if supplier_code.startswith('W'):
            water_features_products.append(product)
        # Priority 2: Organic Feature Pots (Supplier_Item_Code starts with "NBP", "FX", "NR", or "RV", or Description contains "Organic Fiberglass")
        elif supplier_code.startswith('NBP') or supplier_code.startswith('FX') or supplier_code.startswith('NR') or supplier_code.startswith('RV') or 'ORGANIC FIBERGLASS' in description:
            nb_fx_products.append(product)
        # Priority 3: Driptrays (Supplier_Item_Code contains "JPRD" or "JPQD")
        elif 'JPRD' in supplier_code or 'JPQD' in supplier_code:
            driptrays_products.append(product)
        # Priority 4: Drop-in Pots (Description contains "Styler" or "Shrub")
        elif 'STYLER' in description or 'SHRUB' in description:
            dropin_products.append(product)
        # Priority 5: Japi Planters (Supplier_Item_Code starts with "J" or "NSGB2")
        elif supplier_code.startswith('J') or supplier_code.startswith('NSGB2'):
            japi_products.append(product)
        # Priority 6: Plastic Woven Mats (Supplier_Item_Code starts with "PWM")
        elif supplier_code.startswith('PWM'):
            plasticmats_products.append(product)
    
    # Create a dictionary mapping supplier codes to on-hand quantities for water feature calculations
    products_dict = {}
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        if supplier_code:
            products_dict[supplier_code] = product.get('Onhand_Available', '0')
    
    # Calculate on-hand quantities for water features
    for product in water_features_products:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
        product['Onhand_Available'] = str(calculated_onhand)
    
    # Sort all lists alphabetically by Description
    japi_products = sorted(japi_products, key=lambda x: (x.get('Description') or '').lower())
    dropin_products = sorted(dropin_products, key=lambda x: (x.get('Description') or '').lower())
    driptrays_products = sorted(driptrays_products, key=lambda x: (x.get('Description') or '').lower())
    water_features_products = sorted(water_features_products, key=lambda x: (x.get('Description') or '').lower())
    nb_fx_products = sorted(nb_fx_products, key=lambda x: (x.get('Description') or '').lower())
    plasticmats_products = sorted(plasticmats_products, key=lambda x: (x.get('Description') or '').lower())
    
    # Filter all_products to exclude products with certain words in description
    excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
    all_products_filtered = []
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip()
        # Skip products without a supplier item code
        if not supplier_code:
            continue
        description = (product.get('Description') or '').strip()
        # Check if description contains any excluded words (case-insensitive)
        if not any(word.lower() in description.lower() for word in excluded_words):
            all_products_filtered.append(product)
    
    # Calculate on-hand quantities for water features in all_products
    for product in all_products_filtered:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        if supplier_code.startswith('W'):
            calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
            product['Onhand_Available'] = str(calculated_onhand)
    
    all_products = sorted(all_products_filtered, key=lambda x: (x.get('Description') or '').lower())
    
    return render_template('index.html', error=None, japi_products=japi_products, dropin_products=dropin_products, driptrays_products=driptrays_products, water_features_products=water_features_products, nb_fx_products=nb_fx_products, plasticmats_products=plasticmats_products, all_products=all_products, cache_timestamp=cache_timestamp, is_cached=is_cached, api_error=api_error, is_staff=True)


@app.route('/export')
def export_excel():
    """Export stock data to Excel file - exports products from the specified tab"""
    from flask import request
    
    # Get the tab parameter from query string
    tab = request.args.get('tab', 'allproducts')
    
    # Fetch all stock data from cache
    result, cache_timestamp, is_cached, _ = get_cached_stock_data()
    
    # Handle errors
    if isinstance(result, dict) and 'error' in result:
        return f"Error: {result['error']}", 400
    
    if not isinstance(result, list):
        result = []
    
    # Create a dictionary mapping supplier codes to on-hand quantities for water feature calculations
    products_dict = {}
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        if supplier_code:
            products_dict[supplier_code] = product.get('Onhand_Available', '0')
    
    # Filter products based on the tab (same logic as index route)
    if tab == 'allproducts':
        # Export all products (excluding products with certain words in description)
        excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
        products = []
        for product in result:
            supplier_code = (product.get('Supplier_Item_Code') or '').strip()
            # Skip products without a supplier item code
            if not supplier_code:
                continue
            description = (product.get('Description') or '').strip()
            # Check if description contains any excluded words (case-insensitive)
            if not any(word.lower() in description.lower() for word in excluded_words):
                products.append(product)
        # Calculate on-hand quantities for water features in all_products
        for product in products:
            supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
            if supplier_code.startswith('W'):
                calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
                product['Onhand_Available'] = str(calculated_onhand)
        sheet_name = 'All Products'
        filename_prefix = 'all_products'
    else:
        # Separate products into categories (same logic as index route)
        water_features_products = []
        nb_fx_products = []
        driptrays_products = []
        dropin_products = []
        japi_products = []
        plasticmats_products = []
        
        for product in result:
            supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
            # Skip products without a supplier item code
            if not supplier_code:
                continue
            description = (product.get('Description') or '').strip().upper()
            
            # Priority 1: Water Features (Supplier_Item_Code starts with "W")
            if supplier_code.startswith('W'):
                water_features_products.append(product)
            # Priority 2: Organic Feature Pots (Supplier_Item_Code starts with "NBP", "FX", "NR", or "RV", or Description contains "Organic Fiberglass")
            elif supplier_code.startswith('NBP') or supplier_code.startswith('FX') or supplier_code.startswith('NR') or supplier_code.startswith('RV') or 'ORGANIC FIBERGLASS' in description:
                nb_fx_products.append(product)
            # Priority 3: Driptrays (Supplier_Item_Code contains "JPRD" or "JPQD")
            elif 'JPRD' in supplier_code or 'JPQD' in supplier_code:
                driptrays_products.append(product)
            # Priority 4: Drop-in Pots (Description contains "Styler" or "Shrub")
            elif 'STYLER' in description or 'SHRUB' in description:
                dropin_products.append(product)
            # Priority 5: Japi Planters (Supplier_Item_Code starts with "J" or "NSGB2")
            elif supplier_code.startswith('J') or supplier_code.startswith('NSGB2'):
                japi_products.append(product)
            # Priority 6: Plastic Woven Mats (Supplier_Item_Code starts with "PWM")
            elif supplier_code.startswith('PWM'):
                plasticmats_products.append(product)
        
        # Calculate on-hand quantities for water features
        for product in water_features_products:
            supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
            calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
            product['Onhand_Available'] = str(calculated_onhand)
        
        # Get products for the requested tab
        if tab == 'japi':
            products = japi_products
            sheet_name = 'Japi Planters'
            filename_prefix = 'japi_planters'
        elif tab == 'driptrays':
            products = driptrays_products
            sheet_name = 'Driptrays'
            filename_prefix = 'driptrays'
        elif tab == 'dropin':
            products = dropin_products
            sheet_name = 'Drop-in Pots'
            filename_prefix = 'dropin_pots'
        elif tab == 'waterfeatures':
            products = water_features_products
            sheet_name = 'Water Features'
            filename_prefix = 'water_features'
        elif tab == 'nbfx':
            products = nb_fx_products
            sheet_name = 'Organic Feature Pots'
            filename_prefix = 'organic_feature_pots'
        elif tab == 'plasticmats':
            products = plasticmats_products
            sheet_name = 'Plastic Woven Mats'
            filename_prefix = 'plastic_woven_mats'
        else:
            # Default to all products if tab is unknown (excluding products with certain words in description)
            excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
            products = []
            for product in result:
                supplier_code = (product.get('Supplier_Item_Code') or '').strip()
                # Skip products without a supplier item code
                if not supplier_code:
                    continue
                description = (product.get('Description') or '').strip()
                # Check if description contains any excluded words (case-insensitive)
                if not any(word.lower() in description.lower() for word in excluded_words):
                    products.append(product)
            # Calculate on-hand quantities for water features
            for product in products:
                supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
                if supplier_code.startswith('W'):
                    calculated_onhand = calculate_water_feature_onhand(supplier_code, products_dict)
                    product['Onhand_Available'] = str(calculated_onhand)
            sheet_name = 'All Products'
            filename_prefix = 'all_products'
    
    # Sort products alphabetically by Description
    products = sorted(products, key=lambda x: (x.get('Description') or '').lower())
    
    # Create DataFrame
    df = pd.DataFrame(products)
    
    # All tabs include Description, Supplier_Item_Code, and Onhand_Available
    required_columns = ['Description', 'Supplier_Item_Code', 'Onhand_Available']
    
    # Ensure columns exist
    for col in required_columns:
        if col not in df.columns:
            df[col] = ''
    
    # Select only required columns
    df = df[required_columns]
    
    # Create Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    
    output.seek(0)
    
    # Generate filename with timestamp
    timestamp = datetime.now(UTC_PLUS_2).strftime('%Y%m%d_%H%M%S')
    filename = f'{filename_prefix}_{timestamp}.xlsx'
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    # Initialize cache on startup
    print("Initializing cache...")
    get_cached_stock_data()
    
    # Start background thread to update cache periodically
    cache_thread = threading.Thread(target=update_cache_periodically, daemon=True)
    cache_thread.start()
    print("Cache update thread started. Cache will refresh every 5 minutes.")
    
    app.run(debug=True, host='127.0.0.1', port=6200)

