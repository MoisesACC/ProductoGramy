"use client"

import type React from "react"

import { useState } from "react"
import Link from "next/link"
import { useRouter } from "next/navigation"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import Image from "next/image"
import { registerUser } from "@/lib/auth"
import { Navbar } from "@/components/navbar"

export default function RegistroPage() {
  const router = useRouter()
  const [formData, setFormData] = useState({
    nombre: "",
    email: "",
    password: "",
    confirmPassword: "",
  })
  const [errors, setErrors] = useState<Record<string, string>>({})

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()

    const newErrors: Record<string, string> = {}

    // Validation
    if (!formData.nombre) newErrors.nombre = "El nombre es requerido"
    if (!formData.email) newErrors.email = "El correo es requerido"
    if (!formData.password) newErrors.password = "La contraseña es requerida"
    if (formData.password !== formData.confirmPassword) {
      newErrors.confirmPassword = "Las contraseñas no coinciden"
    }

    if (Object.keys(newErrors).length > 0) {
      setErrors(newErrors)
      return
    }

    const success = registerUser(formData.email, formData.nombre, formData.password)

    if (success) {
      // Redirect to login
      router.push("/login")
    } else {
      setErrors({ email: "Este correo ya está registrado" })
    }
  }

  const handleChange = (field: string, value: string) => {
    setFormData((prev) => ({ ...prev, [field]: value }))
    setErrors((prev) => ({ ...prev, [field]: "" }))
  }

  return (
    <div>
      <Navbar />
      <div className="min-h-screen bg-[#DBCFAE] flex items-center justify-center p-4">
        <div className="w-full max-w-md">
          <div className="bg-[#EDE4CC] rounded-lg shadow-lg p-8">
            {/* Logo */}
            <div className="flex justify-center mb-4 md:mb-6">
              <div className="w-40 h-40 md:w-100 md:h-72 relative">
                <Image
                  src="/logo.jpg"
                  alt="Postres Originales Logo"
                  fill
                  className="object-contain"
                />
              </div>
            </div>

            {/* Welcome Message */}
            <h1 className="text-2xl font-bold text-center text-[#3E2723] mb-8">Crear una cuenta</h1>

            {/* Registration Form */}
            <form onSubmit={handleSubmit} className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="nombre" className="text-[#3E2723]">
                  Nombre Completo
                </Label>
                <Input
                  id="nombre"
                  type="text"
                  placeholder="Eva Perez Sobrevilla"
                  value={formData.nombre}
                  onChange={(e) => handleChange("nombre", e.target.value)}
                  className={`bg-white border-[#D7CCC8] ${errors.nombre ? "border-red-500" : ""}`}
                />
                {errors.nombre && <p className="text-xs text-red-600">{errors.nombre}</p>}
              </div>

              <div className="space-y-2">
                <Label htmlFor="email" className="text-[#3E2723]">
                  Correo Electrónico
                </Label>
                <Input
                  id="email"
                  type="email"
                  placeholder="eva.perez@gmail.com"
                  value={formData.email}
                  onChange={(e) => handleChange("email", e.target.value)}
                  className={`bg-white border-[#D7CCC8] ${errors.email ? "border-red-500" : ""}`}
                />
                {errors.email && <p className="text-xs text-red-600">{errors.email}</p>}
              </div>

              <div className="space-y-2">
                <Label htmlFor="password" className="text-[#3E2723]">
                  Contraseña
                </Label>
                <Input
                  id="password"
                  type="password"
                  placeholder="••••••••"
                  value={formData.password}
                  onChange={(e) => handleChange("password", e.target.value)}
                  className={`bg-white border-[#D7CCC8] ${errors.password ? "border-red-500" : ""}`}
                />
                {errors.password && <p className="text-xs text-red-600">{errors.password}</p>}
              </div>

              <div className="space-y-2">
                <Label htmlFor="confirmPassword" className="text-[#3E2723]">
                  Confirmar Contraseña
                </Label>
                <Input
                  id="confirmPassword"
                  type="password"
                  placeholder="••••••••"
                  value={formData.confirmPassword}
                  onChange={(e) => handleChange("confirmPassword", e.target.value)}
                  className={`bg-white border-[#D7CCC8] ${errors.confirmPassword ? "border-red-500" : ""}`}
                />
                {errors.confirmPassword && <p className="text-xs text-red-600">{errors.confirmPassword}</p>}
              </div>

              <Button type="submit" className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white">
                Registrarse
              </Button>
            </form>

            {/* Login Link */}
            <p className="text-center text-sm text-[#6D4C41] mt-6">
              ¿Ya tienes una cuenta?{" "}
              <Link href="/login" className="text-[#4F5729] font-medium hover:underline">
                Iniciar sesión
              </Link>
            </p>
          </div>
        </div>
      </div>
    </div>
  )
}
