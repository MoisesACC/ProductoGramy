import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Truck, Clock, MapPin, Phone } from "lucide-react"

export default function DeliveryPage() {
  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <h1 className="text-4xl font-bold text-[#3E2723] mb-8">Información de Delivery</h1>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          <Card className="bg-white">
            <CardContent className="p-6">
              <div className="flex items-start gap-4">
                <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center flex-shrink-0">
                  <Truck className="h-6 w-6 text-[#4F5729]" />
                </div>
                <div>
                  <h3 className="font-semibold text-[#3E2723] mb-2">Envío Rápido</h3>
                  <p className="text-sm text-[#6D4C41]">
                    Realizamos entregas en un plazo de 30-45 minutos dentro de la zona de cobertura
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>

          <Card className="bg-white">
            <CardContent className="p-6">
              <div className="flex items-start gap-4">
                <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center flex-shrink-0">
                  <Clock className="h-6 w-6 text-[#4F5729]" />
                </div>
                <div>
                  <h3 className="font-semibold text-[#3E2723] mb-2">Horario de Atención</h3>
                  <p className="text-sm text-[#6D4C41]">Lunes a Domingo: 8:00 AM - 8:00 PM</p>
                </div>
              </div>
            </CardContent>
          </Card>

          <Card className="bg-white">
            <CardContent className="p-6">
              <div className="flex items-start gap-4">
                <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center flex-shrink-0">
                  <MapPin className="h-6 w-6 text-[#4F5729]" />
                </div>
                <div>
                  <h3 className="font-semibold text-[#3E2723] mb-2">Zona de Cobertura</h3>
                  <p className="text-sm text-[#6D4C41]">Cubrimos toda la zona metropolitana. Costo de envío: S/ 5.00</p>
                </div>
              </div>
            </CardContent>
          </Card>

          <Card className="bg-white">
            <CardContent className="p-6">
              <div className="flex items-start gap-4">
                <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center flex-shrink-0">
                  <Phone className="h-6 w-6 text-[#4F5729]" />
                </div>
                <div>
                  <h3 className="font-semibold text-[#3E2723] mb-2">Contacto</h3>
                  <p className="text-sm text-[#6D4C41]">WhatsApp: +51 999 999 999</p>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>

        <Card className="bg-white">
          <CardContent className="p-6">
            <h2 className="text-2xl font-bold text-[#3E2723] mb-4">Política de Envío</h2>
            <ul className="space-y-2 text-[#6D4C41]">
              <li>• Pedido mínimo: S/ 20.00</li>
              <li>• Envío gratis en compras mayores a S/ 50.00</li>
              <li>• Recibirás una notificación cuando tu pedido esté en camino</li>
              <li>• Puedes rastrear tu pedido en tiempo real</li>
              <li>• Aceptamos pagos en efectivo y transferencia</li>
            </ul>
          </CardContent>
        </Card>
      </main>
    </div>
  )
}
