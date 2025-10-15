"use client"

import Link from "next/link"
import { Search, ShoppingCart, Menu, X, LogOut, LayoutDashboard } from "lucide-react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { useCart } from "@/components/cart-provider"
import { useAuth } from "@/components/auth-provider"
import { useState } from "react"
import { useRouter } from "next/navigation"

export function Navbar() {
  const { items } = useCart()
  const { user, isAuthenticated, logout } = useAuth()
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false)
  const router = useRouter()

  const cartItemCount = items.reduce((sum, item) => sum + item.quantity, 0)

  const handleLogout = () => {
    logout()
    router.push("/")
  }

  return (
    <nav className="sticky top-0 z-50 w-full border-b border-border/40 bg-[#DBCFAE] backdrop-blur supports-[backdrop-filter]:bg-[#DBCFAE]/95">
      <div className="container mx-auto flex h-16 items-center justify-between px-4">
        <Link href="/" className="flex items-center gap-2">
          <div className="text-xl md:text-2xl font-bold text-[#4F5729]">🍞 GRAMY</div>
        </Link>

        <div className="hidden items-center gap-4 lg:gap-6 lg:flex">
          <Link href="/" className="text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors">
            Inicio
          </Link>
          <Link href="/productos" className="text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors">
            Productos
          </Link>
          <Link
            href="/promociones"
            className="text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors"
          >
            Promociones
          </Link>
          <Link href="/delivery" className="text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors">
            Delivery
          </Link>
          <Link href="/contacto" className="text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors">
            Contacto
          </Link>
        </div>

        <div className="flex items-center gap-2 md:gap-3">
          <div className="relative hidden lg:block">
            <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-[#6D4C41]" />
            <Input
              type="search"
              placeholder="Buscar productos..."
              className="w-[200px] pl-8 bg-white border-[#D7CCC8]"
            />
          </div>

          <Link href="/carrito">
            <Button variant="ghost" size="icon" className="relative hover:bg-[#EDE4CC]">
              <ShoppingCart className="h-5 w-5 text-[#3E2723]" />
              {cartItemCount > 0 && (
                <span className="absolute -right-1 -top-1 flex h-5 w-5 items-center justify-center rounded-full bg-[#4F5729] text-xs text-white">
                  {cartItemCount}
                </span>
              )}
            </Button>
          </Link>

          {isAuthenticated && user ? (
            <div className="hidden md:flex items-center gap-2">
              <span className="text-sm font-medium text-[#3E2723]">Hola, {user.name}</span>
              {user.isAdmin && (
                <Link href="/dashboard">
                  <Button variant="ghost" size="icon" className="hover:bg-[#EDE4CC]" title="Dashboard">
                    <LayoutDashboard className="h-5 w-5 text-[#3E2723]" />
                  </Button>
                </Link>
              )}
              <Button
                variant="ghost"
                size="icon"
                onClick={handleLogout}
                className="hover:bg-[#EDE4CC]"
                title="Cerrar sesión"
              >
                <LogOut className="h-5 w-5 text-[#3E2723]" />
              </Button>
            </div>
          ) : (
            <Link href="/login" className="hidden md:block">
              <Button className="bg-[#4F5729] hover:bg-[#3E4520] text-white">Iniciar Sesión</Button>
            </Link>
          )}

          <Button
            variant="ghost"
            size="icon"
            className="lg:hidden hover:bg-[#EDE4CC]"
            onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
          >
            {isMobileMenuOpen ? <X className="h-6 w-6 text-[#3E2723]" /> : <Menu className="h-6 w-6 text-[#3E2723]" />}
          </Button>
        </div>
      </div>

      {isMobileMenuOpen && (
        <div className="lg:hidden border-t border-border/40 bg-[#DBCFAE]">
          <div className="container mx-auto px-4 py-4 space-y-3">
            <Link
              href="/"
              className="block py-2 text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors"
              onClick={() => setIsMobileMenuOpen(false)}
            >
              Inicio
            </Link>
            <Link
              href="/productos"
              className="block py-2 text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors"
              onClick={() => setIsMobileMenuOpen(false)}
            >
              Productos
            </Link>
            <Link
              href="/promociones"
              className="block py-2 text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors"
              onClick={() => setIsMobileMenuOpen(false)}
            >
              Promociones
            </Link>
            <Link
              href="/delivery"
              className="block py-2 text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors"
              onClick={() => setIsMobileMenuOpen(false)}
            >
              Delivery
            </Link>
            <Link
              href="/contacto"
              className="block py-2 text-sm font-medium text-[#3E2723] hover:text-[#4F5729] transition-colors"
              onClick={() => setIsMobileMenuOpen(false)}
            >
              Contacto
            </Link>

            <div className="relative pt-2">
              <Search className="absolute left-2.5 top-4.5 h-4 w-4 text-[#6D4C41]" />
              <Input
                type="search"
                placeholder="Buscar productos..."
                className="w-full pl-8 bg-white border-[#D7CCC8]"
              />
            </div>

            {isAuthenticated && user ? (
              <div className="pt-2 space-y-2">
                <p className="text-sm font-medium text-[#3E2723]">Hola, {user.name}</p>
                {user.isAdmin && (
                  <Link href="/dashboard" onClick={() => setIsMobileMenuOpen(false)}>
                    <Button
                      variant="outline"
                      className="w-full border-[#4F5729] text-[#4F5729] hover:bg-[#4F5729] hover:text-white bg-transparent"
                    >
                      Dashboard
                    </Button>
                  </Link>
                )}
                <Button
                  onClick={() => {
                    handleLogout()
                    setIsMobileMenuOpen(false)
                  }}
                  variant="outline"
                  className="w-full border-[#4F5729] text-[#4F5729] hover:bg-[#4F5729] hover:text-white"
                >
                  Cerrar Sesión
                </Button>
              </div>
            ) : (
              <Link href="/login" className="block pt-2" onClick={() => setIsMobileMenuOpen(false)}>
                <Button className="w-full bg-[#4F5729] hover:bg-[#3E4520] text-white">Iniciar Sesión</Button>
              </Link>
            )}
          </div>
        </div>
      )}
    </nav>
  )
}
