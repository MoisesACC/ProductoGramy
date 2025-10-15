"use client"

import type React from "react"

import { useState } from "react"
import Link from "next/link"
import { useRouter } from "next/navigation"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import Image from "next/image"
import { loginUser } from "@/lib/auth"
import { useAuth } from "@/components/auth-provider"
import { Navbar } from "@/components/navbar"

export default function LoginPage() {
  const router = useRouter()
  const { refreshUser } = useAuth()
  const [email, setEmail] = useState("")
  const [password, setPassword] = useState("")
  const [error, setError] = useState(false)

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()

    const user = loginUser(email, password)

    if (user) {
      // Refresh auth context
      refreshUser()
      // Redirect to home
      router.push("/")
    } else {
      setError(true)
    }
  }

  return (
    <div>
      <Navbar />
      <div className="min-h-screen bg-[#DBCFAE] flex items-center justify-center p-4">
        <div className="w-full max-w-md">
          <div className="bg-[#EDE4CC] rounded-lg shadow-lg p-6 md:p-8">
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
            <h1 className="text-xl md:text-2xl font-bold text-center text-[#3E2723] mb-6 md:mb-8">BIENVENIDOS A GRAMY</h1>

            {/* Login Form */}
            <form onSubmit={handleSubmit} className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="email" className="text-[#3E2723]">
                  Usuario
                </Label>
                <Input
                  id="email"
                  type="email"
                  placeholder="eva.perez@gmail.com"
                  value={email}
                  onChange={(e) => {
                    setEmail(e.target.value)
                    setError(false)
                  }}
                  className={`bg-white border-[#D7CCC8] ${error ? "border-red-500" : ""}`}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="password" className="text-[#3E2723]">
                  Contraseña
                </Label>
                <Input
                  id="password"
                  type="password"
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value)
                    setError(false)
                  }}
                  className={`bg-white border-[#D7CCC8] ${error ? "border-red-500" : ""}`}
                />
              </div>

              {error && (
                <p className="text-sm text-red-600">
                  Usuario o contraseña incorrectos. Por favor, verifica tus credenciales.
                </p>
              )}

              <Button type="submit" className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white">
                Ingresar
              </Button>
            </form>

            {/* Register Link */}
            <p className="text-center text-sm text-[#6D4C41] mt-6">
              ¿No tienes una cuenta?{" "}
              <Link href="/registro" className="text-[#4F5729] font-medium hover:underline">
                Crear Cuenta
              </Link>
            </p>
          </div>
        </div>
      </div>
    </div>
  )
}
