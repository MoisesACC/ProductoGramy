"use client"

import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { useCart } from "@/components/cart-provider"
import { Trash2, Plus, Minus } from "lucide-react"
import Image from "next/image"
import Link from "next/link"

export default function CarritoPage() {
  const { items, removeItem, updateQuantity, total } = useCart()
  const shippingCost = 5.0
  const subtotal = total
  const totalWithShipping = subtotal + shippingCost

  if (items.length === 0) {
    return (
      <div className="min-h-screen bg-[#EDE4CC]">
        <Navbar />
        <main className="container mx-auto px-4 py-6 md:py-8">
          <div className="max-w-4xl mx-auto">
            <div className="bg-[#4F5729] text-white p-4 md:p-6 rounded-t-lg">
              <h1 className="text-2xl md:text-3xl font-bold">Carrito de Compras</h1>
            </div>
            <Card className="rounded-t-none">
              <CardContent className="p-8 md:p-12 text-center">
                <p className="text-lg md:text-xl text-[#6D4C41] mb-6">Tu carrito está vacío</p>
                <Link href="/productos">
                  <Button className="bg-[#4F5729] hover:bg-[#3E4520] text-white">Ver Productos</Button>
                </Link>
              </CardContent>
            </Card>
          </div>
        </main>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-6 md:py-8">
        <div className="max-w-5xl mx-auto">
          <div className="bg-[#4F5729] text-white p-4 md:p-6 rounded-t-lg">
            <h1 className="text-2xl md:text-3xl font-bold">Carrito de Compras</h1>
          </div>

          <Card className="rounded-t-none">
            <CardContent className="p-0">
              {/* Table Header - Desktop Only */}
              <div className="hidden lg:grid grid-cols-12 gap-4 p-6 bg-[#DBCFAE] border-b border-[#D7CCC8]">
                <div className="col-span-5 font-semibold text-[#3E2723]">Producto</div>
                <div className="col-span-2 font-semibold text-[#3E2723] text-center">Cantidad</div>
                <div className="col-span-2 font-semibold text-[#3E2723] text-center">Precio Unitario</div>
                <div className="col-span-2 font-semibold text-[#3E2723] text-center">Total</div>
                <div className="col-span-1"></div>
              </div>

              {/* Cart Items */}
              <div className="divide-y divide-[#D7CCC8]">
                {items.map((item) => (
                  <div
                    key={item.id}
                    className="grid grid-cols-1 lg:grid-cols-12 gap-3 md:gap-4 p-4 md:p-6 items-center"
                  >
                    {/* Product Info */}
                    <div className="lg:col-span-5 flex items-center gap-3 md:gap-4">
                      <div className="relative w-16 h-16 md:w-20 md:h-20 flex-shrink-0">
                        <Image
                          src={item.image || "/placeholder.svg"}
                          alt={item.name}
                          fill
                          className="object-cover rounded"
                        />
                      </div>
                      <div className="flex-1">
                        <h3 className="font-semibold text-[#3E2723] text-sm md:text-base">{item.name}</h3>
                        {/* Mobile: Show price below name */}
                        <p className="lg:hidden text-sm text-[#6D4C41] mt-1">S/ {item.price.toFixed(2)}</p>
                      </div>
                    </div>

                    {/* Quantity Controls */}
                    <div className="lg:col-span-2 flex items-center justify-between lg:justify-center gap-2">
                      <span className="lg:hidden text-sm font-medium text-[#3E2723]">Cantidad:</span>
                      <div className="flex items-center gap-2">
                        <Button
                          variant="outline"
                          size="icon"
                          className="h-8 w-8 bg-transparent"
                          onClick={() => updateQuantity(item.id, item.quantity - 1)}
                        >
                          <Minus className="h-4 w-4" />
                        </Button>
                        <span className="w-12 text-center font-medium text-[#3E2723]">{item.quantity}</span>
                        <Button
                          variant="outline"
                          size="icon"
                          className="h-8 w-8 bg-transparent"
                          onClick={() => updateQuantity(item.id, item.quantity + 1)}
                        >
                          <Plus className="h-4 w-4" />
                        </Button>
                      </div>
                    </div>

                    {/* Unit Price - Desktop Only */}
                    <div className="hidden lg:block lg:col-span-2 text-center">
                      <p className="text-[#3E2723]">S/ {item.price.toFixed(2)}</p>
                    </div>

                    {/* Total Price */}
                    <div className="lg:col-span-2 flex justify-between lg:justify-center items-center">
                      <span className="lg:hidden text-sm font-medium text-[#3E2723]">Total:</span>
                      <p className="font-semibold text-[#4F5729]">S/ {(item.price * item.quantity).toFixed(2)}</p>
                    </div>

                    {/* Delete Button */}
                    <div className="lg:col-span-1 flex justify-end lg:justify-center">
                      <Button
                        variant="ghost"
                        size="icon"
                        className="text-[#4F5729] hover:text-red-600 hover:bg-red-50"
                        onClick={() => removeItem(item.id)}
                      >
                        <Trash2 className="h-5 w-5" />
                      </Button>
                    </div>
                  </div>
                ))}
              </div>

              {/* Summary */}
              <div className="p-4 md:p-6 bg-[#DBCFAE] space-y-2 md:space-y-3">
                <div className="flex justify-between text-[#3E2723] text-sm md:text-base">
                  <span>Subtotal:</span>
                  <span className="font-medium">S/ {subtotal.toFixed(2)}</span>
                </div>
                <div className="flex justify-between text-[#3E2723] text-sm md:text-base">
                  <span>Envío:</span>
                  <span className="font-medium">S/ {shippingCost.toFixed(2)}</span>
                </div>
                <div className="flex justify-between text-lg md:text-xl font-bold text-[#3E2723] pt-2 md:pt-3 border-t border-[#D7CCC8]">
                  <span>Total:</span>
                  <span>S/ {totalWithShipping.toFixed(2)}</span>
                </div>
              </div>

              {/* Action Buttons */}
              <div className="p-4 md:p-6 flex justify-end">
                <Link href="/pago" className="w-full md:w-auto">
                  <Button className="w-full md:w-auto bg-[#4F5729] hover:bg-[#3E4520] text-white px-6 md:px-8">
                    Proceder al Pago
                  </Button>
                </Link>
              </div>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  )
}
