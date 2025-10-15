export interface Product {
  id: string
  name: string
  description: string
  price: number
  image: string
  category: "mazamorra" | "harina" | "bebida" | "galleta"
}

export const products: Product[] = [
  {
    id: "1",
    name: "Mazamorra de Tocosh",
    description: "Postre tradicional de tocosh, rico y nutritivo",
    price: 15.0,
    image: "/mazamorra.png",
    category: "mazamorra",
  },
  {
    id: "2",
    name: "Harina de Tocosh",
    description: "Harina de tocosh, ideal para preparaciones saludables",
    price: 18.0,
    image: "/HarinaTocosh.png",
    category: "harina",
  },
  {
    id: "3",
    name: "Bebida de Tocosh",
    description: "Refrescante bebida de tocosh, muy buena para la salud",
    price: 22.0,
    image: "/tocosh-drink-bottle (2).jpg",
    category: "bebida",
  },
  {
    id: "4",
    name: "Galleta de Tocosh",
    description: "Deliciosas galletas de tocosh, perfectas para cualquier momento",
    price: 10.0,
    image: "/tocosh-cookies.jpg",
    category: "galleta",
  }
]

export function getStoredProducts(): Product[] {
  if (typeof window === "undefined") return products
  const stored = localStorage.getItem("gramy_products")
  return stored ? JSON.parse(stored) : products
}

export function saveProduct(product: Omit<Product, "id">): Product {
  const allProducts = getStoredProducts()
  const newProduct: Product = {
    ...product,
    id: Date.now().toString(),
  }
  const updatedProducts = [...allProducts, newProduct]
  localStorage.setItem("gramy_products", JSON.stringify(updatedProducts))
  return newProduct
}

export function updateProduct(id: string, updates: Partial<Product>): void {
  const allProducts = getStoredProducts()
  const updatedProducts = allProducts.map((p) => (p.id === id ? { ...p, ...updates } : p))
  localStorage.setItem("gramy_products", JSON.stringify(updatedProducts))
}

export function deleteProduct(id: string): void {
  const allProducts = getStoredProducts()
  const updatedProducts = allProducts.filter((p) => p.id !== id)
  localStorage.setItem("gramy_products", JSON.stringify(updatedProducts))
}
