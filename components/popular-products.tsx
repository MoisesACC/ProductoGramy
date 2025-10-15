"use client"

import { Card, CardContent, CardFooter } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { products } from "@/lib/products"
import { useCart } from "@/components/cart-provider"
import { useToast } from "@/hooks/use-toast"
import Image from "next/image"

export function PopularProducts() {
  const { addItem } = useCart()
  const { toast } = useToast()

  const handleAddToCart = (product: (typeof products)[0]) => {
    addItem({
      id: product.id,
      name: product.name,
      price: product.price,
      image: product.image,
    })
    toast({
      title: "Producto agregado",
      description: `${product.name} se agregó al carrito`,
    })
  }

  return (
    <section>
      <h2 className="text-3xl font-bold text-[#3E2723] mb-6">Productos Populares</h2>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        {products.map((product) => (
          <Card key={product.id} className="overflow-hidden hover:shadow-lg transition-shadow">
            <CardContent className="p-0">
              <div className="relative h-48">
                <Image src={product.image || "/placeholder.svg"} alt={product.name} fill className="object-cover" />
              </div>
              <div className="p-4">
                <h3 className="font-semibold text-[#3E2723] mb-2">{product.name}</h3>
                <p className="text-sm text-[#6D4C41] mb-3">{product.description}</p>
                <p className="text-xl font-bold text-[#4F5729]">S/ {product.price.toFixed(2)}</p>
              </div>
            </CardContent>
            <CardFooter className="p-4 pt-0">
              <Button
                className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white"
                onClick={() => handleAddToCart(product)}
              >
                Agregar al Carrito
              </Button>
            </CardFooter>
          </Card>
        ))}
      </div>
    </section>
  )
}
