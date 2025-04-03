import pandas as pd
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from py3dbp import Packer, Bin, Item
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

def read_excel(file_path, sheet_name='Sheet1'):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df

def get_order_product_dimensions(order_sku, products):
    sku = str(order_sku).lstrip('0')
    product = products[products['SKU'] == sku]
    if not product.empty:
        try:
            length = Decimal(product.iloc[0]['Length']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            width = Decimal(product.iloc[0]['Width']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            height = Decimal(product.iloc[0]['Height']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            weight = Decimal(product.iloc[0]['Weight Unit JD']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            return (length, width, height), weight
        except InvalidOperation:
            return None, None
    else:
        return None, None

def aggregate_order_dimensions(order, products):
    items = []
    for sku in order['SKU']:
        product_dimensions, product_weight = get_order_product_dimensions(sku, products)
        if product_dimensions is not None and product_weight is not None:
            item = Item(sku, *product_dimensions, product_weight)
            items.append(item)
    return items

def find_optimal_carton(items, cartons):
    packer = Packer()

    for _, carton in cartons.iterrows():
        try:
            carton_dimensions = (Decimal(carton['Length']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), 
                                 Decimal(carton['Width']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), 
                                 Decimal(carton['Height']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
            bin = Bin(carton['Code'], *carton_dimensions, Decimal('9999999.99'))
            bin.cost = Decimal(carton['Price']).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            packer.add_bin(bin)
        except InvalidOperation:
            pass

    for item in items:
        packer.add_item(item)

    packer.pack()

    optimal_bin = None
    min_cost = float('inf')

    for b in packer.bins:
        if b.unfitted_items:
            continue
        total_cost = b.cost
        if total_cost < min_cost:
            min_cost = total_cost
            optimal_bin = b

    return optimal_bin, packer.bins

def visualize_packing(optimal_bin):
    fig = plt.figure()
    ax = fig.add_subplot(111, projection='3d')
    
    x = [0, float(optimal_bin.width), float(optimal_bin.width), 0, 0, float(optimal_bin.width), float(optimal_bin.width), 0]
    y = [0, 0, float(optimal_bin.height), float(optimal_bin.height), 0, 0, float(optimal_bin.height), float(optimal_bin.height)]
    z = [0, 0, 0, 0, float(optimal_bin.depth), float(optimal_bin.depth), float(optimal_bin.depth), float(optimal_bin.depth)]
    
    verts = [[x[0], y[0], z[0]], [x[1], y[1], z[1]], [x[2], y[2], z[2]], [x[3], y[3], z[3]], [x[4], y[4], z[4]], [x[5], y[5], z[5]], [x[6], y[6], z[6]], [x[7], y[7], z[7]]]
    faces = [[verts[j] for j in [0, 1, 5, 4]], [verts[j] for j in [1, 2, 6, 5]], [verts[j] for j in [2, 3, 7, 6]], [verts[j] for j in [3, 0, 4, 7]], [verts[j] for j in [0, 1, 2, 3]], [verts[j] for j in [4, 5, 6, 7]]]
    ax.add_collection3d(Poly3DCollection(faces, facecolors='cyan', linewidths=1, edgecolors='r', alpha=.25))

    colors = ['blue', 'green', 'red', 'yellow', 'purple', 'orange', 'pink', 'brown']
    for idx, item in enumerate(optimal_bin.items):
        x = [float(item.position[0]), float(item.position[0] + item.width), float(item.position[0] + item.width), float(item.position[0]), float(item.position[0]), float(item.position[0] + item.width), float(item.position[0] + item.width), float(item.position[0])]
        y = [float(item.position[1]), float(item.position[1]), float(item.position[1] + item.height), float(item.position[1] + item.height), float(item.position[1]), float(item.position[1]), float(item.position[1] + item.height), float(item.position[1] + item.height)]
        z = [float(item.position[2]), float(item.position[2]), float(item.position[2]), float(item.position[2]), float(item.position[2] + item.depth), float(item.position[2] + item.depth), float(item.position[2] + item.depth), float(item.position[2] + item.depth)]

        verts = [[x[0], y[0], z[0]], [x[1], y[1], z[1]], [x[2], y[2], z[2]], [x[3], y[3], z[3]], [x[4], y[4], z[4]], [x[5], y[5], z[5]], [x[6], y[6], z[6]], [x[7], y[7], z[7]]]
        faces = [[verts[j] for j in [0, 1, 5, 4]], [verts[j] for j in [1, 2, 6, 5]], [verts[j] for j in [2, 3, 7, 6]], [verts[j] for j in [3, 0, 4, 7]], [verts[j] for j in [0, 1, 2, 3]], [verts[j] for j in [4, 5, 6, 7]]]
        ax.add_collection3d(Poly3DCollection(faces, facecolors=colors[idx % len(colors)], linewidths=1, edgecolors='r', alpha=.5))

        ax.text(float(item.position[0] + item.width / 2), float(item.position[1] + item.height / 2), float(item.position[2] + item.depth / 2), item.name, color='black')

    ax.set_xlim(0, float(optimal_bin.width))
    ax.set_ylim(0, float(optimal_bin.height))
    ax.set_zlim(0, float(optimal_bin.depth))

    ax.set_xlabel('Length')
    ax.set_ylabel('Width')
    ax.set_zlabel('Height')
    ax.grid(True)

    plt.show()

def main():
    orders_file_path = r'C:\Users\pwisniewski\Desktop\IT\py\PackOptimizer\test.xlsx'
    cartons_file_path = r'C:\Users\pwisniewski\Desktop\IT\py\PackOptimizer\Dane.xlsx'
    products_file_path = r'C:\Users\pwisniewski\Desktop\IT\py\PackOptimizer\Dane.xlsx'
    cartons_sheet_name = 'Karton'
    
    orders = read_excel(orders_file_path)
    orders['SKU'] = orders['SKU'].astype(str)
    cartons = read_excel(cartons_file_path, cartons_sheet_name)
    products = read_excel(products_file_path)
    products['SKU'] = products['SKU'].astype(str)
    
    grouped_orders = orders.groupby('Order')
    
    for order_id, order_group in grouped_orders:
        items = aggregate_order_dimensions(order_group, products)
        
        if items:
            optimal_bin, bins = find_optimal_carton(items, cartons)
            
            if optimal_bin is not None:
                print(f"Zamówienie: {order_id}")
                print(f"Optymalny karton: {optimal_bin.name}")
                print(f"Wymiary kartonu: {optimal_bin.width}x{optimal_bin.height}x{optimal_bin.depth}")
                print(f"Koszt kartonu: {optimal_bin.cost}")
                
                print(f"Sposób ułożenia towarów w kartonie {optimal_bin.name}:")
                for item in optimal_bin.items:
                    print(f" - {item.name}: ({item.width}x{item.height}x{item.depth}) waga: {item.weight}")
                
                visualize_packing(optimal_bin)
            else:
                print(f"Zamówienie: {order_id}")
                print("Nie znaleziono odpowiedniego kartonu")
                for item in items:
                    print(f"Produkt {item.name} ({item.width}x{item.height}x{item.depth}) nie zmieścił się w żadnym z dostępnych kartonów:")
                    for _, carton in cartons.iterrows():
                        carton_dimensions = (carton['Length'], carton['Width'], carton['Height'])
                        print(f" - {carton['Code']}: {carton_dimensions} koszt: {carton['Price']}")
        else:
            print(f"Nie znaleziono danych produktu dla zamówienia: {order_id}")

if __name__ == "__main__":
    main()