"use client"

import type React from "react"

import { useState, useEffect } from "react"
import { useRouter } from "next/navigation"
import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group"
import { Checkbox } from "@/components/ui/checkbox"
import { useCart } from "@/components/cart-provider"
import { useAuth } from "@/components/auth-provider"
import { createOrder } from "@/lib/orders"
import { Upload, AlertCircle } from "lucide-react"

export default function PagoPage() {
  const router = useRouter()
  const { total, items } = useCart()
  const { isAuthenticated, user } = useAuth()
  const [paymentMethod, setPaymentMethod] = useState<"card" | "yape">("card")
  const [acceptedTerms, setAcceptedTerms] = useState(false)
  const [cardData, setCardData] = useState({
    number: "",
    expiry: "",
    cvv: "",
    name: "",
  })
  const [yapeScreenshot, setYapeScreenshot] = useState<File | null>(null)

  useEffect(() => {
    if (!isAuthenticated) {
      router.push("/login")
    }
  }, [isAuthenticated, router])

  if (items.length === 0) {
    return (
      <div className="min-h-screen bg-[#EDE4CC]">
        <Navbar />
        <main className="container mx-auto px-4 py-8">
          <div className="max-w-2xl mx-auto">
            <Card>
              <CardContent className="p-8 text-center">
                <AlertCircle className="h-16 w-16 text-[#6D4C41] mx-auto mb-4" />
                <h2 className="text-2xl font-bold text-[#3E2723] mb-2">Tu carrito está vacío</h2>
                <p className="text-[#6D4C41] mb-6">Agrega productos a tu carrito antes de proceder al pago.</p>
                <Button
                  onClick={() => router.push("/productos")}
                  className="bg-[#4F5729] hover:bg-[#3E4520] text-white"
                >
                  Ver Productos
                </Button>
              </CardContent>
            </Card>
          </div>
        </main>
      </div>
    )
  }

  if (!isAuthenticated) {
    return null
  }

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()

    if (!acceptedTerms) {
      alert("Debes aceptar los términos y condiciones")
      return
    }

    if (paymentMethod === "card") {
      if (!cardData.number || !cardData.expiry || !cardData.cvv || !cardData.name) {
        alert("Por favor completa todos los campos de la tarjeta")
        return
      }
    } else if (paymentMethod === "yape") {
      if (!yapeScreenshot) {
        alert("Por favor sube la captura de pago de Yape")
        return
      }
    }

    const shippingCost = 5.0
    const totalWithShipping = total + shippingCost

    const order = createOrder(
      user!.email, // userId
      user!.email,
      user!.name,
      items,
      totalWithShipping,
      paymentMethod,
    )

    router.push(`/confirmacion?order=${order.id}`)
  }

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setYapeScreenshot(e.target.files[0])
    }
  }

  const shippingCost = 5.0
  const totalWithShipping = total + shippingCost

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <div className="max-w-2xl mx-auto">
          <div className="bg-[#4F5729] text-white p-6 rounded-t-lg">
            <h1 className="text-3xl font-bold">Seleccione Método de Pago</h1>
          </div>

          <Card className="rounded-t-none">
            <CardContent className="p-6">
              <form onSubmit={handleSubmit} className="space-y-6">
                {/* Payment Method Selection */}
                <RadioGroup value={paymentMethod} onValueChange={(value) => setPaymentMethod(value as "card" | "yape")}>
                  <div className="flex items-center space-x-2 p-4 border border-[#D7CCC8] rounded-lg cursor-pointer hover:bg-[#DBCFAE]/30">
                    <RadioGroupItem value="card" id="card" />
                    <Label htmlFor="card" className="cursor-pointer flex-1 text-[#3E2723]">
                      Tarjeta de Crédito / Débito
                    </Label>
                  </div>

                  <div className="flex items-center space-x-2 p-4 border border-[#D7CCC8] rounded-lg cursor-pointer hover:bg-[#DBCFAE]/30">
                    <RadioGroupItem value="yape" id="yape" />
                    <Label htmlFor="yape" className="cursor-pointer flex-1 text-[#3E2723]">
                      Billetera Digital (Yape)
                    </Label>
                  </div>
                </RadioGroup>

                {/* Card Payment Form */}
                {paymentMethod === "card" && (
                  <div className="space-y-4 p-6 bg-[#DBCFAE]/30 rounded-lg">
                    <h3 className="font-semibold text-[#3E2723] mb-4">Detalles de la tarjeta</h3>

                    <div className="space-y-2">
                      <Label htmlFor="cardNumber" className="text-[#3E2723]">
                        Número de Tarjeta
                      </Label>
                      <Input
                        id="cardNumber"
                        placeholder="Introduce el número de tarjeta"
                        value={cardData.number}
                        onChange={(e) => setCardData({ ...cardData, number: e.target.value })}
                        className="bg-white border-[#D7CCC8]"
                        maxLength={16}
                      />
                    </div>

                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-2">
                        <Label htmlFor="expiry" className="text-[#3E2723]">
                          Fecha de Vencimiento
                        </Label>
                        <Input
                          id="expiry"
                          placeholder="MM/YY"
                          value={cardData.expiry}
                          onChange={(e) => setCardData({ ...cardData, expiry: e.target.value })}
                          className="bg-white border-[#D7CCC8]"
                          maxLength={5}
                        />
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="cvv" className="text-[#3E2723]">
                          CVV
                        </Label>
                        <Input
                          id="cvv"
                          placeholder="123"
                          value={cardData.cvv}
                          onChange={(e) => setCardData({ ...cardData, cvv: e.target.value })}
                          className="bg-white border-[#D7CCC8]"
                          maxLength={3}
                        />
                      </div>
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="cardName" className="text-[#3E2723]">
                        Nombre del titular de la tarjeta
                      </Label>
                      <Input
                        id="cardName"
                        placeholder="Ingrese el nombre del titular de la tarjeta"
                        value={cardData.name}
                        onChange={(e) => setCardData({ ...cardData, name: e.target.value })}
                        className="bg-white border-[#D7CCC8]"
                      />
                    </div>
                  </div>
                )}

                {/* Yape Payment Form */}
                {paymentMethod === "yape" && (
                  <div className="space-y-4 p-6 bg-[#DBCFAE]/30 rounded-lg">
                    <h3 className="font-semibold text-[#3E2723] mb-4">Pago Yape</h3>

                    <div className="space-y-4">
                      <p className="text-[#3E2723]">Para pagar con Yape sigue estos pasos:</p>

                      <ol className="list-decimal list-inside space-y-2 text-[#6D4C41]">
                        <li>Envía el importe total al +51999999999</li>
                        <li>Sube la captura de pantalla del pago a continuación.</li>
                      </ol>

                      <div className="p-4 bg-white border-2 border-dashed border-[#D7CCC8] rounded-lg text-center">
                        <Label htmlFor="screenshot" className="cursor-pointer">
                          <div className="flex flex-col items-center gap-2">
                            <Upload className="h-8 w-8 text-[#6D4C41]" />
                            <span className="text-[#3E2723] font-medium">Cargar captura de pantalla</span>
                            <span className="text-sm text-[#6D4C41]">
                              {yapeScreenshot ? yapeScreenshot.name : "Arrastra y suelta o haga clic para cargar"}
                            </span>
                          </div>
                          <Input
                            id="screenshot"
                            type="file"
                            accept="image/*"
                            className="hidden"
                            onChange={handleFileChange}
                          />
                        </Label>
                      </div>
                    </div>
                  </div>
                )}

                {/* Terms and Conditions */}
                <div className="flex items-start space-x-2">
                  <Checkbox
                    id="terms"
                    checked={acceptedTerms}
                    onCheckedChange={(checked) => setAcceptedTerms(checked as boolean)}
                  />
                  <Label htmlFor="terms" className="text-sm text-[#3E2723] cursor-pointer leading-relaxed">
                    Aceptar Términos y Condiciones
                  </Label>
                </div>

                {/* Order Summary */}
                <div className="p-4 bg-[#DBCFAE]/50 rounded-lg space-y-2">
                  <div className="flex justify-between text-[#3E2723]">
                    <span>Subtotal:</span>
                    <span>S/ {total.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between text-[#3E2723]">
                    <span>Envío:</span>
                    <span>S/ {shippingCost.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between text-xl font-bold text-[#3E2723] pt-2 border-t border-[#D7CCC8]">
                    <span>Total:</span>
                    <span>S/ {totalWithShipping.toFixed(2)}</span>
                  </div>
                </div>

                {/* Submit Button */}
                <Button type="submit" className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white py-6 text-lg">
                  Confirmar Pago
                </Button>
              </form>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  )
}
