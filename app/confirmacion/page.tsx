"use client"

import { useEffect, useState, useRef, Suspense } from "react"
import { useSearchParams } from "next/navigation"
import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { useCart } from "@/components/cart-provider"
import { CheckCircle2 } from "lucide-react"
import Link from "next/link"
import Image from "next/image"

function ConfirmacionContent() {
  const searchParams = useSearchParams()
  const orderNumber = searchParams.get("order")
  const { items, total, clearCart } = useCart()
  const [orderDate] = useState(new Date().toLocaleDateString("es-PE"))
  const [orderItems, setOrderItems] = useState<typeof items>([])
  const [orderTotal, setOrderTotal] = useState(0)
  const hasCleared = useRef(false)

  useEffect(() => {
    // Only run once when component mounts with a valid order
    if (orderNumber && !hasCleared.current && items.length > 0) {
      // Save order details before clearing
      setOrderItems([...items])
      setOrderTotal(total)
      // Clear the cart
      clearCart()
      // Mark as cleared
      hasCleared.current = true
    }
  }, []) // Empty dependency array - only run on mount

  const shippingCost = 5.0
  const totalWithShipping = orderTotal + shippingCost

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <div className="max-w-3xl mx-auto">
          <div className="bg-[#4F5729] text-white p-6 rounded-t-lg">
            <div className="flex items-center gap-3">
              <CheckCircle2 className="h-8 w-8" />
              <h1 className="text-3xl font-bold">¡Gracias por su pedido!</h1>
            </div>
          </div>

          <Card className="rounded-t-none">
            <CardContent className="p-6 space-y-6">
              {/* Success Message */}
              <div className="text-center py-4">
                <p className="text-lg text-[#3E2723]">
                  Su pedido se ha realizado correctamente. Agradecemos su preferencia y esperamos que disfrute de su
                  comida.
                </p>
              </div>

              {/* Order Summary Header */}
              <div>
                <h2 className="text-2xl font-bold text-[#3E2723] mb-4">Resumen del pedido</h2>

                <div className="grid grid-cols-2 gap-4 p-4 bg-[#DBCFAE]/30 rounded-lg">
                  <div>
                    <p className="text-sm text-[#6D4C41]">Order Number</p>
                    <p className="font-semibold text-[#3E2723]">#{orderNumber}</p>
                  </div>
                  <div>
                    <p className="text-sm text-[#6D4C41]">Date</p>
                    <p className="font-semibold text-[#3E2723]">{orderDate}</p>
                  </div>
                </div>
              </div>

              {/* Order Details */}
              <div>
                <h3 className="font-semibold text-[#3E2723] mb-4">Detalles del pedido</h3>
                <div className="space-y-4">
                  {orderItems.map((item) => (
                    <div
                      key={item.id}
                      className="flex items-center gap-4 p-4 bg-white rounded-lg border border-[#D7CCC8]"
                    >
                      <div className="relative w-16 h-16 flex-shrink-0">
                        <Image
                          src={item.image || "/placeholder.svg"}
                          alt={item.name}
                          fill
                          className="object-cover rounded"
                        />
                      </div>
                      <div className="flex-1">
                        <h4 className="font-semibold text-[#3E2723]">{item.name}</h4>
                        <p className="text-sm text-[#6D4C41]">Cantidad: {item.quantity}</p>
                      </div>
                      <div className="text-right">
                        <p className="font-semibold text-[#4F5729]">S/ {(item.price * item.quantity).toFixed(2)}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Total Summary */}
              <div className="p-4 bg-[#DBCFAE]/50 rounded-lg space-y-2">
                <div className="flex justify-between text-[#3E2723]">
                  <span>Total Paid</span>
                  <span className="font-semibold">S/ {totalWithShipping.toFixed(2)}</span>
                </div>
              </div>

              {/* Next Steps */}
              <div className="p-6 bg-[#DBCFAE]/30 rounded-lg">
                <h3 className="font-semibold text-[#3E2723] mb-2">Próximos pasos</h3>
                <p className="text-[#6D4C41]">
                  Tu pedido estará listo para envío en 30 minutos. Recibirás una notificación cuando esté listo.
                </p>
              </div>

              {/* Action Buttons */}
              <div className="flex flex-col sm:flex-row gap-4">
                <Link href="/historial" className="flex-1">
                  <Button
                    variant="outline"
                    className="w-full border-[#4F5729] text-[#4F5729] hover:bg-[#4F5729] hover:text-white bg-transparent"
                  >
                    Ver historial de pedidos
                  </Button>
                </Link>
                <Link href="/" className="flex-1">
                  <Button className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white">Volver al Inicio</Button>
                </Link>
              </div>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  )
}

export default function ConfirmacionPage() {
  return (
    <Suspense fallback={<div>Loading...</div>}>
      <ConfirmacionContent />
    </Suspense>
  )
}
