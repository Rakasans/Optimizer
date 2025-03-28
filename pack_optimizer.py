import pandas as pd
# import logging
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from py3dbp import Packer, Bin, Item
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# Konfiguracja loggera
# logging.basicConfig(filename='pack_optimizer.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
# logger = logging.getLogger()

def read_excel(file_path, sheet_name='Sheet1'):
    # Czytanie danych z pliku Excel, z określonego arkusza
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # logger.info(f'Załadowano dane z pliku: {file_path}, arkusz: {sheet_name}')
    return df

def get_order_product_dimensions(order_sku, products):
    # Usunięcie 12 zer z przodu numeru SKU
    sku = str(order_sku).lstrip('0')
    
    # Znalezienie wymiarów produktu na podstawie SKU
    product = products[products['SKU'] == sku]
    if not product.empty:
        try:
            length = Decimal(product.iloc[0]['Length']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)  # Upewnij się, że nazwa kolumny to "Length"
            width = Decimal(product.iloc[0]['Width']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            height = Decimal(product.iloc[0]['Height']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            weight = Decimal(product.iloc[0]['Weight Unit JD']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            # logger.info(f"SKU: {sku} -> Wymiary: ({length}, {width}, {height}), Waga: {weight}")
            return (length, width, height), weight
        except InvalidOperation as e:
            # logger.error(f"InvalidOperation error for SKU: {sku} -> {e}")
            return None, None
    else:
        # logger.warning(f"SKU: {sku} -> Nie znaleziono danych")
        return None, None

def aggregate_order_dimensions(order, products):
    items = []
    for sku in order['SKU']:
        product_dimensions, product_weight = get_order_product_dimensions(sku, products)
        if product_dimensions is not None and product_weight is not None:
            item = Item(sku, *product_dimensions, product_weight)
            items.append(item)
        else:
            # logger.warning(f"Nie znaleziono danych produktu dla SKU: {sku}")
            pass
    # logger.info(f"Zagregowane wymiary zamówienia: {order['Order'].iloc[0]}")
    return items

def find_optimal_carton(items, cartons):
    packer = Packer()

    for _, carton in cartons.iterrows():
        try:
            carton_dimensions = (Decimal(carton['Length']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), 
                                 Decimal(carton['Width']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), 
                                 Decimal(carton['Height']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))  # Upewnij się, że nazwa kolumny to "Length"
            bin = Bin(carton['Code'], *carton_dimensions, Decimal('9999999.99'))
            bin.cost = Decimal(carton['Price']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)  # Custom attribute to store the cost
            packer.add_bin(bin)
            # logger.info(f"Dodano karton: {carton['Code']} -> Wymiary: {carton_dimensions}, Koszt: {carton['Price']}")
        except InvalidOperation as e:
            # logger.error(f"InvalidOperation error for Carton: {carton['Code']} -> {e}")
            pass

    for item in items:
        # logger.info(f"Dodawanie produktu: {item.name} ({item.width}x{item.height}x{item.depth}) waga: {item.weight}")
        packer.add_item(item)

    packer.pack()

    optimal_bin = None
    min_cost = float('inf')

    for b in packer.bins:
        # logger.info(f"Sprawdzanie kartonu: {b.name}")
        if b.unfitted_items:
            for unfitted_item in b.unfitted_items:
                # logger.warning(f"Nie dopasowano produktu: {unfitted_item.name} ({unfitted_item.width}x{unfitted_item.height}x{unfitted_item.depth})")
                # logger.debug(f"Pozostała pojemność kartonu {b.name}: {b.width * b.height * b.depth - sum([item.width * item.height * item.depth for item in b.items])}")
                pass
            continue
        total_cost = b.cost
        if total_cost < min_cost:
            min_cost = total_cost
            optimal_bin = b

    return optimal_bin, packer.bins

def visualize_packing(optimal_bin):
    fig = plt.figure()
    ax = fig.add_subplot(111, projection='3d')
    
    # Rysowanie kartonu
    x = [0, float(optimal_bin.width), float(optimal_bin.width), 0, 0, float(optimal_bin.width), float(optimal_bin.width), 0]
    y = [0, 0, float(optimal_bin.height), float(optimal_bin.height), 0, 0, float(optimal_bin.height), float(optimal_bin.height)]
    z = [0, 0, 0, 0, float(optimal_bin.depth), float(optimal_bin.depth), float(optimal_bin.depth), float(optimal_bin.depth)]
    
    verts = [[x[0], y[0], z[0]], [x[1], y[1], z[1]], [x[2], y[2], z[2]], [x[3], y[3], z[3]], [x[4], y[4], z[4]], [x[5], y[5], z[5]], [x[6], y[6], z[6]], [x[7], y[7], z[7]]]
    faces = [[verts[j] for j in [0, 1, 5, 4]], [verts[j] for j in [1, 2, 6, 5]], [verts[j] for j in [2, 3, 7, 6]], [verts[j] for j in [3, 0, 4, 7]], [verts[j] for j in [0, 1, 2, 3]], [verts[j] for j in [4, 5, 6, 7]]]
    ax.add_collection3d(Poly3DCollection(faces, facecolors='cyan', linewidths=1, edgecolors='r', alpha=.25))

    # Rysowanie produktów w kartonie
    for item in optimal_bin.items:
        x = [float(item.position[0]), float(item.position[0] + item.width), float(item.position[0] + item.width), float(item.position[0]), float(item.position[0]), float(item.position[0] + item.width), float(item.position[0] + item.width), float(item.position[0])]
        y = [float(item.position[1]), float(item.position[1]), float(item.position[1] + item.height), float(item.position[1] + item.height), float(item.position[1]), float(item.position[1]), float(item.position[1] + item.height), float(item.position[1] + item.height)]
        z = [float(item.position[2]), float(item.position[2]), float(item.position[2]), float(item.position[2]), float(item.position[2] + item.depth), float(item.position[2] + item.depth), float(item.position[2] + item.depth), float(item.position[2] + item.depth)]
        
        verts = [[x[0], y[0], z[0]], [x[1], y[1], z[1]], [x[2], y[2], z[2]], [x[3], y[3], z[3]], [x[4], y[4], z[4]], [x[5], y[5], z[5]], [x[6], y[6], z[6]], [x[7], y[7], z[7]]]
        faces = [[verts[j] for j in [0, 1, 5, 4]], [verts[j] for j in [1, 2, 6, 5]], [verts[j] for j in [2, 3, 7, 6]], [verts[j] for j in [3, 0, 4, 7]], [verts[j] for j in [0, 1, 2, 3]], [verts[j] for j in [4, 5, 6, 7]]]
        ax.add_collection3d(Poly3DCollection(faces, facecolors='blue', linewidths=1, edgecolors='r', alpha=.5))
    
    ax.set_xlabel('Length')
    ax.set_ylabel('Width')
    ax.set_zlabel('Height')
    plt.show()

def main():
    # Ścieżki do plików Excel
    orders_file_path = r'C:\Users\pwisniewski\Desktop\IT\py\PackOptimizer\test.xlsx'
    cartons_file_path = r'C:\Users\pwisniewski\Desktop\IT\py\PackOptimizer\Dane.xlsx'
    products_file_path = r'C:\Users\pwisniewski\Desktop\IT\py\PackOptimizer\Dane.xlsx'
    cartons_sheet_name = 'Karton'
    
    # Czytanie danych z plików Excel i konwersja kolumny SKU na typ str
    orders = read_excel(orders_file_path)
    orders['SKU'] = orders['SKU'].astype(str)
    cartons = read_excel(cartons_file_path, cartons_sheet_name)
    products = read_excel(products_file_path)
    products['SKU'] = products['SKU'].astype(str)
    
    # Wyświetlenie załadowanych danych zamówień
    # logger.info("Załadowane dane zamówień:")
    # logger.info(orders)
    
    # Wyświetlenie nagłówków kolumn z pliku Dane.xlsx
    # logger.info("Nagłówki kolumn w pliku Dane.xlsx:")
    # logger.info(products.columns)
    
    # Wyświetlenie kilku pierwszych wierszy z pliku Dane.xlsx
    # logger.info("Pierwsze kilka wierszy z pliku Dane.xlsx:")
    # logger.info(products.head())
    
    # Wyświetlenie dostępnych kartonów
    # logger.info("Dostępne kartony:")
    for _, carton in cartons.iterrows():
        # logger.info(f" - {carton['Code']}: {carton['Length']}x{carton['Width']}x{carton['Height']} koszt: {carton['Price']}")
        pass
    
    # Agregacja wymiarów i wagi dla wszystkich produktów w każdym zamówieniu
    grouped_orders = orders.groupby('Order')
    
    for order_id, order_group in grouped_orders:
        # logger.info(f"\nZamówienie: {order_id}")
        
        items = aggregate_order_dimensions(order_group, products)
        
        if items:
            # logger.info(f"Wymiary produktów zamówienia {order_id}:")
            for item in items:
                # logger.info(f" - {item.name}: {item.width}x{item.height}x{item.depth} waga: {item.weight}")
                pass
            
            optimal_bin, bins = find_optimal_carton(items, cartons)
            
            if optimal_bin is not None:
                # logger.info(f"Optymalny karton dla zamówienia {order_id}: {optimal_bin.name}")
                # logger.info(f"Wymiary kartonu: {optimal_bin.width}x{optimal_bin.height}x{optimal_bin.depth}")
                # logger.info(f"Koszt kartonu: {optimal_bin.cost}")
                print(f"Zamówienie: {order_id}")
                print(f"Optymalny karton: {optimal_bin.name}")
                print(f"Wymiary kartonu: {optimal_bin.width}x{optimal_bin.height}x{optimal_bin.depth}")
                print(f"Koszt kartonu: {optimal_bin.cost}")
                
                # Wyświetlenie sposobu ułożenia towarów
                print(f"Sposób ułożenia towarów w kartonie {optimal_bin.name}:")
                for item in optimal_bin.items:
                    print(f" - {item.name}: ({item.width}x{item.height}x{item.depth}) waga: {item.weight}")
                
                visualize_packing(optimal_bin)
            else:
                # logger.warning(f"Nie znaleziono odpowiedniego kartonu dla zamówienia {order_id}")
                print(f"Zamówienie: {order_id}")
                print("Nie znaleziono odpowiedniego kartonu")
                for item in items:
                    # logger.warning(f"Produkt {item.name} ({item.width}x{item.height}x{item.depth}) nie zmieścił się w żadnym z dostępnych kartonów:")
                    print(f"Produkt {item.name} ({item.width}x{item.height}x{item.depth}) nie zmieścił się w żadnym z dostępnych kartonów:")
                    for _, carton in cartons.iterrows():
                        carton_dimensions = (carton['Length'], carton['Width'], carton['Height'])
                        # logger.warning(f" - {carton['Code']}: {carton_dimensions} koszt: {carton['Price']}")
                        print(f" - {carton['Code']}: {carton_dimensions} koszt: {carton['Price']}")
        else:
            # logger.warning(f"Nie znaleziono danych produktu dla zamówienia: {order_id}")
            print(f"Nie znaleziono danych produktu dla zamówienia: {order_id}")

if __name__ == "__main__":
    main()