"use client"

import { useState, useEffect } from "react"
import { useRouter } from "next/navigation"
import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { useAuth } from "@/components/auth-provider"
import { useCart } from "@/components/cart-provider"
import { getUserOrders } from "@/lib/orders"
import type { Order } from "@/lib/orders"
import { Search, ShoppingCart } from "lucide-react"

export default function HistorialPage() {
  const router = useRouter()
  const { isAuthenticated, user } = useAuth()
  const { addItem } = useCart()
  const [searchOrder, setSearchOrder] = useState("")
  const [orders, setOrders] = useState<Order[]>([])

  useEffect(() => {
    if (!isAuthenticated) {
      router.push("/login")
      return
    }

    if (user) {
      const userOrders = getUserOrders(user.email)
      setOrders(userOrders)
    }
  }, [isAuthenticated, user, router])

  if (!isAuthenticated) {
    return null
  }

  const filteredOrders = orders.filter((order) => order.id.includes(searchOrder))

  const handleReorder = (order: Order) => {
    order.items.forEach((item) => {
      for (let i = 0; i < item.quantity; i++) {
        addItem({
          id: item.id,
          name: item.name,
          price: item.price,
          image: item.image,
        })
      }
    })
    router.push("/carrito")
  }

  const getStatusColor = (status: string) => {
    switch (status) {
      case "Pendiente":
        return "bg-orange-500"
      case "En preparación":
        return "bg-yellow-500"
      case "Listo para envío":
        return "bg-blue-500"
      case "Completado":
        return "bg-green-500"
      case "Rechazado":
        return "bg-red-500"
      default:
        return "bg-gray-500"
    }
  }

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto">
          <h1 className="text-4xl font-bold text-[#3E2723] mb-8">Historial de Pedidos</h1>

          <Card className="mb-6 bg-white">
            <CardContent className="p-6">
              <div className="space-y-2">
                <Label htmlFor="orderSearch" className="text-[#3E2723]">
                  Buscar por número de orden
                </Label>
                <div className="relative">
                  <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-[#6D4C41]" />
                  <Input
                    id="orderSearch"
                    placeholder="Ingresa el número de orden..."
                    value={searchOrder}
                    onChange={(e) => setSearchOrder(e.target.value)}
                    className="pl-8 bg-white border-[#D7CCC8]"
                  />
                </div>
              </div>
            </CardContent>
          </Card>

          <div className="space-y-4">
            {filteredOrders.length > 0 ? (
              filteredOrders.map((order) => (
                <Card key={order.id} className="bg-white hover:shadow-lg transition-shadow">
                  <CardContent className="p-6">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4">
                      <div>
                        <h3 className="text-xl font-bold text-[#3E2723]">Orden #{order.id}</h3>
                        <p className="text-sm text-[#6D4C41]">{order.date}</p>
                      </div>
                      <Badge className={`${getStatusColor(order.status)} text-white`}>{order.status}</Badge>
                    </div>

                    {order.rejectionReason && (
                      <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg">
                        <p className="text-sm text-red-800">
                          <strong>Motivo de rechazo:</strong> {order.rejectionReason}
                        </p>
                      </div>
                    )}

                    <div className="space-y-2 mb-4">
                      {order.items.map((item, index) => (
                        <div key={index} className="flex justify-between text-sm">
                          <span className="text-[#3E2723]">
                            {item.name} x{item.quantity}
                          </span>
                          <span className="text-[#6D4C41]">S/ {(item.price * item.quantity).toFixed(2)}</span>
                        </div>
                      ))}
                    </div>

                    <div className="space-y-2 pt-4 border-t border-[#D7CCC8]">
                      <div className="flex justify-between text-sm">
                        <span className="text-[#6D4C41]">Subtotal:</span>
                        <span className="text-[#3E2723]">S/ {(order.subtotal || 0).toFixed(2)}</span>
                      </div>
                      <div className="flex justify-between text-sm">
                        <span className="text-[#6D4C41]">Envío:</span>
                        <span className="text-[#3E2723]">S/ {(order.envio || 0).toFixed(2)}</span>
                      </div>
                      <div className="flex justify-between items-center pt-2 border-t border-[#D7CCC8]">
                        <span className="font-semibold text-[#3E2723]">Total:</span>
                        <span className="text-xl font-bold text-[#4F5729]">S/ {(order.total || 0).toFixed(2)}</span>
                      </div>
                      <Button
                        onClick={() => handleReorder(order)}
                        className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white mt-3"
                      >
                        <ShoppingCart className="h-4 w-4 mr-2" />
                        Volver a Pedir
                      </Button>
                    </div>
                  </CardContent>
                </Card>
              ))
            ) : (
              <Card className="bg-white">
                <CardContent className="p-12 text-center">
                  <p className="text-[#6D4C41]">
                    {searchOrder ? "No se encontraron pedidos" : "Aún no has realizado ningún pedido"}
                  </p>
                </CardContent>
              </Card>
            )}
          </div>
        </div>
      </main>
    </div>
  )
}
