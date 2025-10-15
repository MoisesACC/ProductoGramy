"use client"

import { useState, useEffect } from "react"
import { Navbar } from "@/components/navbar"
import { Card, CardContent, CardFooter } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group"
import { getStoredProducts } from "@/lib/products"
import { useCart } from "@/components/cart-provider"
import { useToast } from "@/hooks/use-toast"
import { Search } from "lucide-react"
import Image from "next/image"

export default function ProductosPage() {
  const [searchQuery, setSearchQuery] = useState("")
  const [selectedCategory, setSelectedCategory] = useState<string>("all")
  const [priceRange, setPriceRange] = useState<[number, number]>([0, 100])
  const [products, setProducts] = useState(getStoredProducts())
  const { addItem } = useCart()
  const { toast } = useToast()

  useEffect(() => {
    setProducts(getStoredProducts())
  }, [])

  const filteredProducts = products.filter((product) => {
    const matchesSearch = product.name.toLowerCase().includes(searchQuery.toLowerCase())
    const matchesCategory = selectedCategory === "all" || product.category === selectedCategory
    const matchesPrice = product.price >= priceRange[0] && product.price <= priceRange[1]
    return matchesSearch && matchesCategory && matchesPrice
  })

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
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-6 md:py-8">
        <h1 className="text-3xl md:text-4xl font-bold text-[#3E2723] mb-6 md:mb-8">Nuestros Productos GRAMY</h1>

        <div className="flex flex-col lg:flex-row gap-6 md:gap-8">
          {/* Sidebar Filters */}
          <aside className="w-full lg:w-64 space-y-6">
            <Card className="bg-white">
              <CardContent className="p-4 md:p-6 space-y-4 md:space-y-6">
                {/* Search */}
                <div className="space-y-2">
                  <Label className="text-[#3E2723]">Buscar producto</Label>
                  <div className="relative">
                    <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-[#6D4C41]" />
                    <Input
                      placeholder="Buscar..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="pl-8 bg-white border-[#D7CCC8]"
                    />
                  </div>
                </div>

                {/* Categories */}
                <div className="space-y-2">
                  <Label className="text-[#3E2723] font-semibold">Categorías</Label>
                  <RadioGroup value={selectedCategory} onValueChange={setSelectedCategory}>
                    <div className="flex items-center space-x-2">
                      <RadioGroupItem value="all" id="all" />
                      <Label htmlFor="all" className="cursor-pointer">
                        Todos
                      </Label>
                    </div>
                    <div className="flex items-center space-x-2">
                      <RadioGroupItem value="mazamorra" id="mazamorra" />
                      <Label htmlFor="mazamorra" className="cursor-pointer">
                        Mazamorra
                      </Label>
                    </div>
                    <div className="flex items-center space-x-2">
                      <RadioGroupItem value="harina" id="harina" />
                      <Label htmlFor="harina" className="cursor-pointer">
                        Harina
                      </Label>
                    </div>
                    <div className="flex items-center space-x-2">
                      <RadioGroupItem value="bebida" id="bebida" />
                      <Label htmlFor="bebida" className="cursor-pointer">
                        Bebida
                      </Label>
                    </div>
                    <div className="flex items-center space-x-2">
                      <RadioGroupItem value="galleta" id="galleta" />
                      <Label htmlFor="galleta" className="cursor-pointer">
                        Galleta
                      </Label>
                    </div>
                  </RadioGroup>
                </div>

                {/* Price Range */}
                <div className="space-y-2">
                  <Label className="text-[#3E2723] font-semibold">Rango de Precios</Label>
                  <div className="flex items-center gap-2">
                    <Input
                      type="number"
                      placeholder="Min"
                      value={priceRange[0]}
                      onChange={(e) => setPriceRange([Number(e.target.value), priceRange[1]])}
                      className="bg-white border-[#D7CCC8]"
                    />
                    <span className="text-[#6D4C41]">-</span>
                    <Input
                      type="number"
                      placeholder="Max"
                      value={priceRange[1]}
                      onChange={(e) => setPriceRange([priceRange[0], Number(e.target.value)])}
                      className="bg-white border-[#D7CCC8]"
                    />
                  </div>
                </div>
              </CardContent>
            </Card>
          </aside>

          {/* Products Grid */}
          <div className="flex-1">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-2 xl:grid-cols-3 gap-4 md:gap-6">
              {filteredProducts.map((product) => (
                <Card key={product.id} className="overflow-hidden hover:shadow-lg transition-shadow bg-white">
                  <CardContent className="p-0">
                    <div className="relative h-40 sm:h-44 md:h-48">
                      <Image
                        src={product.image || "/placeholder.svg"}
                        alt={product.name}
                        fill
                        className="object-cover"
                      />
                    </div>
                    <div className="p-3 md:p-4">
                      <h3 className="font-semibold text-[#3E2723] mb-2">{product.name}</h3>
                      <p className="text-sm text-[#6D4C41] mb-3 line-clamp-2">{product.description}</p>
                      <p className="text-lg md:text-xl font-bold text-[#4F5729]">S/ {product.price.toFixed(2)}</p>
                    </div>
                  </CardContent>
                  <CardFooter className="p-3 md:p-4 pt-0">
                    <Button
                      className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white text-sm md:text-base"
                      onClick={() => handleAddToCart(product)}
                    >
                      Agregar al Carrito
                    </Button>
                  </CardFooter>
                </Card>
              ))}
            </div>

            {filteredProducts.length === 0 && (
              <div className="text-center py-12">
                <p className="text-[#6D4C41] text-lg">No se encontraron productos</p>
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  )
}
