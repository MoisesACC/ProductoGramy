"use client"

import type React from "react"

import { useState } from "react"
import { Navbar } from "@/components/navbar"
import { Card, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { Mail, Phone, MapPin } from "lucide-react"

export default function ContactoPage() {
  const [formData, setFormData] = useState({
    nombre: "",
    email: "",
    mensaje: "",
  })

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    // Handle form submission
    alert("Mensaje enviado. Nos pondremos en contacto pronto.")
    setFormData({ nombre: "", email: "", mensaje: "" })
  }

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <h1 className="text-4xl font-bold text-[#3E2723] mb-8">Contáctanos</h1>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {/* Contact Form */}
          <Card className="bg-white">
            <CardContent className="p-6">
              <h2 className="text-2xl font-bold text-[#3E2723] mb-6">Envíanos un mensaje</h2>
              <form onSubmit={handleSubmit} className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="nombre" className="text-[#3E2723]">
                    Nombre
                  </Label>
                  <Input
                    id="nombre"
                    value={formData.nombre}
                    onChange={(e) => setFormData({ ...formData, nombre: e.target.value })}
                    className="bg-white border-[#D7CCC8]"
                    required
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="email" className="text-[#3E2723]">
                    Correo Electrónico
                  </Label>
                  <Input
                    id="email"
                    type="email"
                    value={formData.email}
                    onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                    className="bg-white border-[#D7CCC8]"
                    required
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="mensaje" className="text-[#3E2723]">
                    Mensaje
                  </Label>
                  <Textarea
                    id="mensaje"
                    value={formData.mensaje}
                    onChange={(e) => setFormData({ ...formData, mensaje: e.target.value })}
                    className="bg-white border-[#D7CCC8] min-h-[120px]"
                    required
                  />
                </div>

                <Button type="submit" className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white">
                  Enviar Mensaje
                </Button>
              </form>
            </CardContent>
          </Card>

          {/* Contact Information */}
          <div className="space-y-6">
            <Card className="bg-white">
              <CardContent className="p-6">
                <div className="flex items-start gap-4">
                  <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center flex-shrink-0">
                    <Phone className="h-6 w-6 text-[#4F5729]" />
                  </div>
                  <div>
                    <h3 className="font-semibold text-[#3E2723] mb-2">Teléfono</h3>
                    <p className="text-sm text-[#6D4C41]">+51 999 999 999</p>
                    <p className="text-sm text-[#6D4C41]">WhatsApp disponible</p>
                  </div>
                </div>
              </CardContent>
            </Card>

            <Card className="bg-white">
              <CardContent className="p-6">
                <div className="flex items-start gap-4">
                  <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center flex-shrink-0">
                    <Mail className="h-6 w-6 text-[#4F5729]" />
                  </div>
                  <div>
                    <h3 className="font-semibold text-[#3E2723] mb-2">Correo Electrónico</h3>
                    <p className="text-sm text-[#6D4C41]">contacto@gramy.com</p>
                    <p className="text-sm text-[#6D4C41]">ventas@gramy.com</p>
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
                    <h3 className="font-semibold text-[#3E2723] mb-2">Dirección</h3>
                    <p className="text-sm text-[#6D4C41]">Av. Principal 123</p>
                    <p className="text-sm text-[#6D4C41]">Lima, Perú</p>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </main>
    </div>
  )
}
