export interface Promotion {
  id: string
  title: string
  description: string
  originalPrice: number
  discountPrice: number
  discount: number
  image: string
  items: string[] // Product IDs included in the promotion
}

const defaultPromotions: Promotion[] = [
  {
    id: "promo-1",
    title: "Combo Familiar",
    description: "2 Mazamorras + 1 Bebida + 1 Paquete de Galletas",
    originalPrice: 60,
    discountPrice: 45,
    discount: 25,
    image: "/mazamorra-tocosh-dessert.jpg",
    items: ["1", "3", "4"],
  },
  {
    id: "promo-2",
    title: "Pack Saludable",
    description: "1 Harina + 2 Bebidas de Tocosh",
    originalPrice: 62,
    discountPrice: 50,
    discount: 20,
    image: "/tocosh-drink-bottle.jpg",
    items: ["2", "3"],
  },
  {
    id: "promo-3",
    title: "Oferta del Mes",
    description: "3 Paquetes de Galletas de Tocosh",
    originalPrice: 30,
    discountPrice: 25,
    discount: 17,
    image: "/tocosh-cookies.jpg",
    items: ["4"],
  },
]

export function getStoredPromotions(): Promotion[] {
  if (typeof window === "undefined") return defaultPromotions
  const stored = localStorage.getItem("gramy_promotions")
  return stored ? JSON.parse(stored) : defaultPromotions
}

export function savePromotion(promotion: Omit<Promotion, "id">): Promotion {
  const allPromotions = getStoredPromotions()
  const newPromotion: Promotion = {
    ...promotion,
    id: `promo-${Date.now()}`,
  }
  const updatedPromotions = [...allPromotions, newPromotion]
  localStorage.setItem("gramy_promotions", JSON.stringify(updatedPromotions))
  return newPromotion
}

export function updatePromotion(id: string, updates: Partial<Promotion>): void {
  const allPromotions = getStoredPromotions()
  const updatedPromotions = allPromotions.map((p) => (p.id === id ? { ...p, ...updates } : p))
  localStorage.setItem("gramy_promotions", JSON.stringify(updatedPromotions))
}

export function deletePromotion(id: string): void {
  const allPromotions = getStoredPromotions()
  const updatedPromotions = allPromotions.filter((p) => p.id !== id)
  localStorage.setItem("gramy_promotions", JSON.stringify(updatedPromotions))
}
