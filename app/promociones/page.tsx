"use client"

import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Badge } from "@/components/ui/badge"
import Image from "next/image"
import { useCart } from "@/components/cart-provider"
import { getStoredPromotions } from "@/lib/promotions"
import { useState, useEffect } from "react"
import { ShoppingCart, Check } from "lucide-react"
import { useRouter } from "next/navigation"

export default function PromocionesPage() {
  const { addItem } = useCart()
  const router = useRouter()
  const [promotions, setPromotions] = useState(getStoredPromotions())
  const [addedPromos, setAddedPromos] = useState<Set<string>>(new Set())

  useEffect(() => {
    setPromotions(getStoredPromotions())
  }, [])

  const handleAddToCart = (promo: (typeof promotions)[0]) => {
    addItem({
      id: promo.id,
      name: promo.title,
      price: promo.discountPrice,
      image: promo.image,
    })

    setAddedPromos((prev) => new Set(prev).add(promo.id))

    setTimeout(() => {
      setAddedPromos((prev) => {
        const newSet = new Set(prev)
        newSet.delete(promo.id)
        return newSet
      })
    }, 2000)
  }

  const handleBuyNow = (promo: (typeof promotions)[0]) => {
    addItem({
      id: promo.id,
      name: promo.title,
      price: promo.discountPrice,
      image: promo.image,
    })
    router.push("/carrito")
  }

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <h1 className="text-3xl md:text-4xl font-bold text-[#3E2723] mb-8">Promociones Especiales</h1>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {promotions.map((promo) => (
            <Card key={promo.id} className="overflow-hidden hover:shadow-lg transition-shadow bg-white">
              <CardContent className="p-0">
                <div className="relative h-64">
                  <Image src={promo.image || "/placeholder.svg"} alt={promo.title} fill className="object-cover" />
                  <Badge className="absolute top-4 right-4 bg-red-600 text-white text-lg px-3 py-1">
                    -{promo.discount}%
                  </Badge>
                </div>
                <div className="p-6">
                  <h3 className="text-xl font-bold text-[#3E2723] mb-2">{promo.title}</h3>
                  <p className="text-sm text-[#6D4C41] mb-4">{promo.description}</p>
                  <div className="flex items-center gap-3 mb-4">
                    <span className="text-lg text-[#6D4C41] line-through">S/ {promo.originalPrice.toFixed(2)}</span>
                    <span className="text-2xl font-bold text-[#4F5729]">S/ {promo.discountPrice.toFixed(2)}</span>
                  </div>
                  <div className="flex gap-2">
                    <Button
                      onClick={() => handleAddToCart(promo)}
                      variant="outline"
                      className="flex-1 border-[#4F5729] text-[#4F5729] hover:bg-[#4F5729] hover:text-white"
                    >
                      {addedPromos.has(promo.id) ? (
                        <>
                          <Check className="w-4 h-4 mr-2" />
                          Agregado
                        </>
                      ) : (
                        <>
                          <ShoppingCart className="w-4 h-4 mr-2" />
                          Agregar
                        </>
                      )}
                    </Button>
                    <Button
                      onClick={() => handleBuyNow(promo)}
                      className="flex-1 bg-[#4F5729] hover:bg-[#3E4520] text-white"
                    >
                      Comprar Ahora
                    </Button>
                  </div>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>
      </main>
    </div>
  )
}
