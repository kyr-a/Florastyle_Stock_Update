from flask import Flask, render_template, send_file
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

# API Configuration from Postman collection
API_URL = "http://192.168.5.90:8090/IQRetailRestAPI/v1/IQ_API_Request_Stock_Attributes?callformat=xml"
API_HEADERS = {
    "Authorization": "Basic OTk6QVBJVGVzdA==",
    "Content-Type": "application/json"
}
API_BODY = """<IQ_API>
	<IQ_API_Request_Stock>
		<IQ_Company_Number>F01</IQ_Company_Number>
		<IQ_Terminal_Number>1</IQ_Terminal_Number>
		<IQ_User_Number>11</IQ_User_Number>
		<IQ_User_Password>257B95EDADE4F097BD73E88534511D197D3A545B</IQ_User_Password>
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


@app.route('/')
def index():
    """Main page - fetch and display stock data"""
    result = fetch_stock_data()
    
    # If there's an error, return it
    if isinstance(result, dict) and 'error' in result:
        return render_template('index.html', error=result['error'], japi_products=[], dropin_products=[], driptrays_products=[], water_features_products=[], nb_fx_products=[], all_products=[])
    
    # Convert to list if needed
    if not isinstance(result, list):
        result = []
    
    # Separate products into categories:
    # 1. Water Features (Supplier_Item_Code starts with "W")
    # 2. Organic Feature Pots (Supplier_Item_Code starts with "NB" or "FX")
    # 3. Driptrays (Supplier_Item_Code contains "JPRD" or "JPQD")
    # 4. Drop-in Pots (Description contains "Styler" or "Shrub")
    # 5. Japi Planters (Supplier_Item_Code starts with "J" or "NSGB2")
    # 6. All Products (all products from API - for separate tab)
    water_features_products = []
    nb_fx_products = []
    driptrays_products = []
    dropin_products = []
    japi_products = []
    
    for product in result:
        supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
        description = (product.get('Description') or '').strip().upper()
        
        # Priority 1: Water Features (Supplier_Item_Code starts with "W")
        if supplier_code.startswith('W'):
            water_features_products.append(product)
        # Priority 2: Organic Feature Pots (Supplier_Item_Code starts with "NB" or "FX")
        elif supplier_code.startswith('NB') or supplier_code.startswith('FX'):
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
    
    # Sort all lists alphabetically by Description
    japi_products = sorted(japi_products, key=lambda x: (x.get('Description') or '').lower())
    dropin_products = sorted(dropin_products, key=lambda x: (x.get('Description') or '').lower())
    driptrays_products = sorted(driptrays_products, key=lambda x: (x.get('Description') or '').lower())
    water_features_products = sorted(water_features_products, key=lambda x: (x.get('Description') or '').lower())
    nb_fx_products = sorted(nb_fx_products, key=lambda x: (x.get('Description') or '').lower())
    
    # Filter all_products to exclude products with certain words in description
    excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
    all_products_filtered = []
    for product in result:
        description = (product.get('Description') or '').strip()
        # Check if description contains any excluded words (case-insensitive)
        if not any(word.lower() in description.lower() for word in excluded_words):
            all_products_filtered.append(product)
    
    all_products = sorted(all_products_filtered, key=lambda x: (x.get('Description') or '').lower())
    
    return render_template('index.html', error=None, japi_products=japi_products, dropin_products=dropin_products, driptrays_products=driptrays_products, water_features_products=water_features_products, nb_fx_products=nb_fx_products, all_products=all_products)


@app.route('/export')
def export_excel():
    """Export stock data to Excel file - exports products from the specified tab"""
    from flask import request
    
    # Get the tab parameter from query string
    tab = request.args.get('tab', 'allproducts')
    
    # Fetch all stock data
    result = fetch_stock_data()
    
    # Handle errors
    if isinstance(result, dict) and 'error' in result:
        return f"Error: {result['error']}", 400
    
    if not isinstance(result, list):
        result = []
    
    # Filter products based on the tab (same logic as index route)
    if tab == 'allproducts':
        # Export all products (excluding products with certain words in description)
        excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
        products = []
        for product in result:
            description = (product.get('Description') or '').strip()
            # Check if description contains any excluded words (case-insensitive)
            if not any(word.lower() in description.lower() for word in excluded_words):
                products.append(product)
        sheet_name = 'All Products'
        filename_prefix = 'all_products'
    else:
        # Separate products into categories (same logic as index route)
        water_features_products = []
        nb_fx_products = []
        driptrays_products = []
        dropin_products = []
        japi_products = []
        
        for product in result:
            supplier_code = (product.get('Supplier_Item_Code') or '').strip().upper()
            description = (product.get('Description') or '').strip().upper()
            
            # Priority 1: Water Features (Supplier_Item_Code starts with "W")
            if supplier_code.startswith('W'):
                water_features_products.append(product)
            # Priority 2: Organic Feature Pots (Supplier_Item_Code starts with "NB" or "FX")
            elif supplier_code.startswith('NB') or supplier_code.startswith('FX'):
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
            # Water Features doesn't export Onhand_Available
        elif tab == 'nbfx':
            products = nb_fx_products
            sheet_name = 'Organic Feature Pots'
            filename_prefix = 'organic_feature_pots'
        else:
            # Default to all products if tab is unknown (excluding products with certain words in description)
            excluded_words = ['Delivery', 'Voluto', 'Nativa', 'Miscellaneous']
            products = []
            for product in result:
                description = (product.get('Description') or '').strip()
                # Check if description contains any excluded words (case-insensitive)
                if not any(word.lower() in description.lower() for word in excluded_words):
                    products.append(product)
            sheet_name = 'All Products'
            filename_prefix = 'all_products'
    
    # Sort products alphabetically by Description
    products = sorted(products, key=lambda x: (x.get('Description') or '').lower())
    
    # Create DataFrame
    df = pd.DataFrame(products)
    
    # Handle Water Features tab - only Description and Supplier_Item_Code
    if tab == 'waterfeatures':
        required_columns = ['Description', 'Supplier_Item_Code']
    else:
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
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'{filename_prefix}_{timestamp}.xlsx'
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

