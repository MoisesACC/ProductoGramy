"use client"

import type React from "react"

import { useState, useEffect } from "react"
import { useRouter } from "next/navigation"
import { Navbar } from "@/components/navbar"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { Input } from "@/components/ui/input"
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { useAuth } from "@/components/auth-provider"
import { getOrders, updateOrderStatus, getTotalRevenue, getPendingOrdersCount } from "@/lib/orders"
import { getStoredProducts, saveProduct, deleteProduct, updateProduct, type Product } from "@/lib/products"
import {
  getUsers,
  createUser,
  updateUser,
  toggleUserStatus,
  syncUsersFromAuth,
  type Order,
  type User,
  type UserRole,
} from "@/lib/users"
import { getStoredPromotions, savePromotion, deletePromotion, type Promotion } from "@/lib/promotions"
import { generateOrderReport, getUniqueCustomers, exportToPDF, exportToExcel, type ReportFilters } from "@/lib/reports"
import { Package, ShoppingBag, DollarSign, AlertCircle, Plus, Search, Upload, FileDown } from "lucide-react"

export default function DashboardPage() {
  const router = useRouter()
  const { isAuthenticated, user } = useAuth()
  const [orders, setOrders] = useState<Order[]>([])
  const [selectedOrder, setSelectedOrder] = useState<Order | null>(null)
  const [rejectionReason, setRejectionReason] = useState("")
  const [products, setProducts] = useState<Product[]>([])
  const [isAddingProduct, setIsAddingProduct] = useState(false)
  const [editingProduct, setEditingProduct] = useState<Product | null>(null) // Added editingProduct state
  const [newProduct, setNewProduct] = useState({
    name: "",
    description: "",
    price: "",
    category: "mazamorra" as Product["category"],
    image: "",
  })
  const [users, setUsers] = useState<User[]>([])
  const [searchQuery, setSearchQuery] = useState("")
  const [isAddingUser, setIsAddingUser] = useState(false)
  const [editingUser, setEditingUser] = useState<User | null>(null)
  const [newUser, setNewUser] = useState({
    name: "",
    email: "",
    password: "",
    role: "Cliente" as UserRole,
  })
  const [promotions, setPromotions] = useState<Promotion[]>([])
  const [isAddingPromotion, setIsAddingPromotion] = useState(false)
  const [newPromotion, setNewPromotion] = useState({
    title: "",
    description: "",
    originalPrice: "",
    discountPrice: "",
    discount: "",
    image: "",
  })
  const [reportFilters, setReportFilters] = useState<ReportFilters>({
    startDate: "",
    endDate: "",
    status: "Todos",
    customer: "Todos",
    productType: "",
  })
  const [reportResults, setReportResults] = useState<Order[]>([]) // Added reportResults state
  const [orderNumberSearch, setOrderNumberSearch] = useState("") // Added orderNumberSearch state

  useEffect(() => {
    if (!isAuthenticated || !user?.isAdmin) {
      router.push("/")
      return
    }

    syncUsersFromAuth()
    loadOrders()
    loadProducts()
    loadUsers()
    loadPromotions()
  }, [isAuthenticated, user, router])

  const loadOrders = () => {
    const allOrders = getOrders()
    setOrders(allOrders)
  }

  const loadProducts = () => {
    const allProducts = getStoredProducts()
    setProducts(allProducts)
  }

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>, isEdit = false) => {
    const file = e.target.files?.[0]
    if (file) {
      const reader = new FileReader()
      reader.onloadend = () => {
        const base64String = reader.result as string
        if (isEdit && editingProduct) {
          setEditingProduct({ ...editingProduct, image: base64String })
        } else {
          setNewProduct({ ...newProduct, image: base64String })
        }
      }
      reader.readAsDataURL(file)
    }
  }

  const handleAddProduct = () => {
    if (!newProduct.name || !newProduct.description || !newProduct.price) {
      alert("Por favor completa todos los campos")
      return
    }

    const price = Number.parseFloat(newProduct.price)
    if (isNaN(price) || price <= 0) {
      alert("Por favor ingresa un precio válido")
      return
    }

    saveProduct({
      name: newProduct.name,
      description: newProduct.description,
      price,
      category: newProduct.category,
      image: newProduct.image || "/placeholder.svg?height=200&width=200",
    })

    setNewProduct({
      name: "",
      description: "",
      price: "",
      category: "mazamorra",
      image: "",
    })
    setIsAddingProduct(false)
    loadProducts()
  }

  const handleUpdateProduct = () => {
    if (!editingProduct) return

    if (!editingProduct.name || !editingProduct.description || !editingProduct.price) {
      alert("Por favor completa todos los campos")
      return
    }

    // Ensure price is a number before updating
    const price = Number.parseFloat(editingProduct.price as string)
    if (isNaN(price) || price <= 0) {
      alert("Por favor ingresa un precio válido")
      return
    }

    updateProduct(editingProduct.id, {
      name: editingProduct.name,
      description: editingProduct.description,
      price: price, // Use the parsed price
      category: editingProduct.category,
      image: editingProduct.image,
    })

    setEditingProduct(null)
    loadProducts()
  }

  const handleDeleteProduct = (id: string) => {
    if (confirm("¿Estás segura de que deseas eliminar este producto?")) {
      deleteProduct(id)
      loadProducts()
    }
  }

  const handleStatusChange = (orderId: string, newStatus: Order["status"]) => {
    if (newStatus === "Rechazado" && !rejectionReason) {
      alert("Por favor ingresa un motivo de rechazo")
      return
    }

    updateOrderStatus(orderId, newStatus, newStatus === "Rechazado" ? rejectionReason : undefined)
    setRejectionReason("")
    setSelectedOrder(null)
    loadOrders()
  }

  const loadUsers = () => {
    const allUsers = getUsers()
    setUsers(allUsers)
  }

  const loadPromotions = () => {
    const allPromotions = getStoredPromotions()
    setPromotions(allPromotions)
  }

  const handleAddPromotion = () => {
    if (
      !newPromotion.title ||
      !newPromotion.description ||
      !newPromotion.originalPrice ||
      !newPromotion.discountPrice
    ) {
      alert("Por favor completa todos los campos")
      return
    }

    const originalPrice = Number.parseFloat(newPromotion.originalPrice)
    const discountPrice = Number.parseFloat(newPromotion.discountPrice)

    if (isNaN(originalPrice) || isNaN(discountPrice) || originalPrice <= 0 || discountPrice <= 0) {
      alert("Por favor ingresa precios válidos")
      return
    }

    if (discountPrice >= originalPrice) {
      alert("El precio con descuento debe ser menor al precio original")
      return
    }

    const discount = Math.round(((originalPrice - discountPrice) / originalPrice) * 100)

    savePromotion({
      title: newPromotion.title,
      description: newPromotion.description,
      originalPrice,
      discountPrice,
      discount,
      image: newPromotion.image || "/placeholder.svg?height=200&width=200",
      items: [],
    })

    setNewPromotion({
      title: "",
      description: "",
      originalPrice: "",
      discountPrice: "",
      discount: "",
      image: "",
    })
    setIsAddingPromotion(false)
    loadPromotions()
  }

  const handleDeletePromotion = (id: string) => {
    if (confirm("¿Estás segura de que deseas eliminar esta promoción?")) {
      deletePromotion(id)
      loadPromotions()
    }
  }

  const handleAddUser = () => {
    if (!newUser.name || !newUser.email || !newUser.password) {
      alert("Por favor completa todos los campos")
      return
    }

    const isAdmin = newUser.role === "Administrador"
    const success = createUser(newUser.email, newUser.name, newUser.password, newUser.role, isAdmin)

    if (!success) {
      alert("Ya existe un usuario con ese correo electrónico")
      return
    }

    setNewUser({
      name: "",
      email: "",
      password: "",
      role: "Cliente",
    })
    setIsAddingUser(false)
    loadUsers()
  }

  const handleUpdateUser = () => {
    if (!editingUser) return

    updateUser(editingUser.email, {
      name: editingUser.name,
      role: editingUser.role,
      isAdmin: editingUser.role === "Administrador",
    })

    setEditingUser(null)
    loadUsers()
  }

  const handleToggleUserStatus = (email: string) => {
    toggleUserStatus(email)
    loadUsers()
  }

  const handleGenerateReport = () => {
    if (orderNumberSearch) {
      const order = orders.find((o) => o.id === orderNumberSearch)
      setReportResults(order ? [order] : [])
    } else {
      const results = generateOrderReport(reportFilters)
      setReportResults(results)
    }
  }

  const filteredUsers = users.filter(
    (user) =>
      user.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      user.email.toLowerCase().includes(searchQuery.toLowerCase()),
  )

  const getOrdersByStatus = (status: Order["status"]) => {
    return orders.filter((order) => order.status === status)
  }

  if (!isAuthenticated || !user?.isAdmin) {
    return null
  }

  const totalRevenue = getTotalRevenue()
  const pendingOrders = getPendingOrdersCount()
  const totalOrders = orders.length

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

  const getRoleBadgeColor = (role: UserRole) => {
    switch (role) {
      case "Administrador":
        return "bg-[#4F5729] text-white"
      case "Vendedor":
        return "bg-blue-500 text-white"
      case "Cliente":
        return "bg-green-500 text-white"
      default:
        return "bg-gray-500 text-white"
    }
  }

  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-8">
        <div className="mb-8">
          <h1 className="text-4xl font-bold text-[#3E2723] mb-2">Dashboard de Administración</h1>
          <p className="text-[#6D4C41]">Bienvenida, {user.name}</p>
        </div>

        {/* Statistics Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-8">
          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-[#3E2723]">Total de Pedidos</CardTitle>
              <ShoppingBag className="h-4 w-4 text-[#6D4C41]" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-[#4F5729]">{totalOrders}</div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-[#3E2723]">Pedidos Pendientes</CardTitle>
              <AlertCircle className="h-4 w-4 text-orange-500" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-orange-500">{pendingOrders}</div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-[#3E2723]">Productos</CardTitle>
              <Package className="h-4 w-4 text-[#6D4C41]" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-[#4F5729]">{products.length}</div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
              <CardTitle className="text-sm font-medium text-[#3E2723]">Ingresos Totales</CardTitle>
              <DollarSign className="h-4 w-4 text-green-600" />
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-green-600">S/ {(totalRevenue || 0).toFixed(2)}</div>
            </CardContent>
          </Card>
        </div>

        {/* Tabs */}
        <Tabs defaultValue="orders" className="space-y-6">
          <TabsList className="bg-[#DBCFAE]">
            <TabsTrigger value="orders">Pedidos</TabsTrigger>
            <TabsTrigger value="history">Historial</TabsTrigger> {/* Added Order History Tab */}
            <TabsTrigger value="products">Productos</TabsTrigger>
            <TabsTrigger value="promotions">Promociones</TabsTrigger>
            <TabsTrigger value="users">Usuarios</TabsTrigger>
            <TabsTrigger value="reports">Reportes</TabsTrigger> {/* Added Reports Tab */}
          </TabsList>

          {/* Orders Tab */}
          <TabsContent value="orders" className="space-y-4">
            <Card>
              <CardHeader>
                <CardTitle className="text-[#3E2723]">Gestión de Pedidos</CardTitle>
              </CardHeader>
              <CardContent>
                {orders.length > 0 ? (
                  <div className="space-y-4">
                    {orders.map((order) => (
                      <Card key={order.id} className="bg-[#DBCFAE]/30">
                        <CardContent className="p-6">
                          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-4">
                            <div className="space-y-1">
                              <h3 className="text-lg font-bold text-[#3E2723]">Orden #{order.id}</h3>
                              <p className="text-sm text-[#6D4C41]">Cliente: {order.userName}</p>
                              <p className="text-sm text-[#6D4C41]">Email: {order.userEmail}</p>
                              <p className="text-sm text-[#6D4C41]">Fecha: {order.date}</p>
                              <p className="text-sm text-[#6D4C41]">
                                Método de pago: {order.paymentMethod === "card" ? "Tarjeta" : "Yape"}
                              </p>
                            </div>
                            <div className="flex flex-col gap-2">
                              <Badge className={`${getStatusColor(order.status)} text-white`}>{order.status}</Badge>
                              <span className="text-lg font-bold text-[#4F5729]">
                                S/ {(order.total || 0).toFixed(2)}
                              </span>
                            </div>
                          </div>

                          <div className="space-y-2 mb-4">
                            <h4 className="font-semibold text-[#3E2723]">Productos:</h4>
                            {order.items.map((item, index) => (
                              <div key={index} className="flex justify-between text-sm">
                                <span className="text-[#3E2723]">
                                  {item.name} x{item.quantity}
                                </span>
                                <span className="text-[#6D4C41]">
                                  S/ {((item.price || 0) * item.quantity).toFixed(2)}
                                </span>
                              </div>
                            ))}
                            <div className="flex justify-between text-sm pt-2 border-t border-[#D7CCC8]">
                              <span className="text-[#3E2723]">Subtotal:</span>
                              <span className="text-[#6D4C41]">S/ {(order.subtotal || 0).toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between text-sm">
                              <span className="text-[#3E2723]">Envío:</span>
                              <span className="text-[#6D4C41]">S/ {(order.envio || 0).toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between text-sm font-semibold pt-1">
                              <span className="text-[#3E2723]">Total:</span>
                              <span className="text-[#4F5729]">S/ {(order.total || 0).toFixed(2)}</span>
                            </div>
                          </div>

                          {order.status !== "Completado" && order.status !== "Rechazado" && (
                            <div className="space-y-3 pt-4 border-t border-[#D7CCC8]">
                              <Label className="text-[#3E2723]">Actualizar Estado</Label>
                              <Select
                                onValueChange={(value) => {
                                  if (value === "Rechazado") {
                                    setSelectedOrder(order)
                                  } else {
                                    handleStatusChange(order.id, value as Order["status"])
                                  }
                                }}
                              >
                                <SelectTrigger className="bg-white">
                                  <SelectValue placeholder="Seleccionar nuevo estado" />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="En preparación">En preparación</SelectItem>
                                  <SelectItem value="Listo para envío">Listo para envío</SelectItem>
                                  <SelectItem value="Completado">Completado</SelectItem>
                                  <SelectItem value="Rechazado">Rechazado</SelectItem>
                                </SelectContent>
                              </Select>

                              {selectedOrder?.id === order.id && (
                                <div className="space-y-2">
                                  <Label htmlFor="rejectionReason" className="text-[#3E2723]">
                                    Motivo de Rechazo
                                  </Label>
                                  <Textarea
                                    id="rejectionReason"
                                    placeholder="Ingresa el motivo del rechazo..."
                                    value={rejectionReason}
                                    onChange={(e) => setRejectionReason(e.target.value)}
                                    className="bg-white"
                                  />
                                  <div className="flex gap-2">
                                    <Button
                                      onClick={() => handleStatusChange(order.id, "Rechazado")}
                                      className="bg-red-600 hover:bg-red-700 text-white"
                                    >
                                      Confirmar Rechazo
                                    </Button>
                                    <Button
                                      onClick={() => {
                                        setSelectedOrder(null)
                                        setRejectionReason("")
                                      }}
                                      variant="outline"
                                    >
                                      Cancelar
                                    </Button>
                                  </div>
                                </div>
                              )}
                            </div>
                          )}
                        </CardContent>
                      </Card>
                    ))}
                  </div>
                ) : (
                  <div className="text-center py-12">
                    <p className="text-[#6D4C41]">No hay pedidos registrados aún</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="history" className="space-y-4">
            <Card>
              <CardHeader>
                <CardTitle className="text-[#3E2723]">Historial de Pedidos</CardTitle>
              </CardHeader>
              <CardContent>
                <Tabs defaultValue="completados" className="space-y-4">
                  <TabsList className="bg-[#DBCFAE]">
                    <TabsTrigger value="completados">Completados</TabsTrigger>
                    <TabsTrigger value="listo">Listo para envío</TabsTrigger>
                    <TabsTrigger value="preparacion">En preparación</TabsTrigger>
                    <TabsTrigger value="rechazado">Rechazado</TabsTrigger>
                  </TabsList>

                  <TabsContent value="completados">
                    {getOrdersByStatus("Completado").length > 0 ? (
                      <div className="space-y-4">
                        {getOrdersByStatus("Completado").map((order) => (
                          <Card key={order.id} className="bg-white">
                            <CardContent className="p-4">
                              <div className="flex justify-between items-start mb-3">
                                <div>
                                  <h4 className="font-bold text-[#3E2723]">Orden #{order.id}</h4>
                                  <p className="text-sm text-[#6D4C41]">{order.userName}</p>
                                  <p className="text-sm text-[#6D4C41]">{order.date}</p>
                                </div>
                                <Badge className="bg-green-500 text-white">Completado</Badge>
                              </div>
                              <div className="text-sm space-y-1">
                                {order.items.map((item, idx) => (
                                  <div key={idx} className="flex justify-between">
                                    <span>
                                      {item.name} x{item.quantity}
                                    </span>
                                    <span>S/ {(item.price * item.quantity).toFixed(2)}</span>
                                  </div>
                                ))}
                                <div className="pt-2 border-t flex justify-between">
                                  <span>Subtotal:</span>
                                  <span>S/ {(order.subtotal || 0).toFixed(2)}</span>
                                </div>
                                <div className="flex justify-between">
                                  <span>Envío:</span>
                                  <span>S/ {(order.envio || 0).toFixed(2)}</span>
                                </div>
                                <div className="font-semibold flex justify-between">
                                  <span>Total:</span>
                                  <span>S/ {(order.total || 0).toFixed(2)}</span>
                                </div>
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    ) : (
                      <p className="text-center text-[#6D4C41] py-8">No hay pedidos completados</p>
                    )}
                  </TabsContent>

                  <TabsContent value="listo">
                    {getOrdersByStatus("Listo para envío").length > 0 ? (
                      <div className="space-y-4">
                        {getOrdersByStatus("Listo para envío").map((order) => (
                          <Card key={order.id} className="bg-white">
                            <CardContent className="p-4">
                              <div className="flex justify-between items-start mb-3">
                                <div>
                                  <h4 className="font-bold text-[#3E2723]">Orden #{order.id}</h4>
                                  <p className="text-sm text-[#6D4C41]">{order.userName}</p>
                                  <p className="text-sm text-[#6D4C41]">{order.date}</p>
                                </div>
                                <Badge className="bg-blue-500 text-white">Listo para envío</Badge>
                              </div>
                              <div className="text-sm space-y-1">
                                {order.items.map((item, idx) => (
                                  <div key={idx} className="flex justify-between">
                                    <span>
                                      {item.name} x{item.quantity}
                                    </span>
                                    <span>S/ {(item.price * item.quantity).toFixed(2)}</span>
                                  </div>
                                ))}
                                <div className="pt-2 border-t flex justify-between">
                                  <span>Subtotal:</span>
                                  <span>S/ {(order.subtotal || 0).toFixed(2)}</span>
                                </div>
                                <div className="flex justify-between">
                                  <span>Envío:</span>
                                  <span>S/ {(order.envio || 0).toFixed(2)}</span>
                                </div>
                                <div className="font-semibold flex justify-between">
                                  <span>Total:</span>
                                  <span>S/ {(order.total || 0).toFixed(2)}</span>
                                </div>
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    ) : (
                      <p className="text-center text-[#6D4C41] py-8">No hay pedidos listos para envío</p>
                    )}
                  </TabsContent>

                  <TabsContent value="preparacion">
                    {getOrdersByStatus("En preparación").length > 0 ? (
                      <div className="space-y-4">
                        {getOrdersByStatus("En preparación").map((order) => (
                          <Card key={order.id} className="bg-white">
                            <CardContent className="p-4">
                              <div className="flex justify-between items-start mb-3">
                                <div>
                                  <h4 className="font-bold text-[#3E2723]">Orden #{order.id}</h4>
                                  <p className="text-sm text-[#6D4C41]">{order.userName}</p>
                                  <p className="text-sm text-[#6D4C41]">{order.date}</p>
                                </div>
                                <Badge className="bg-yellow-500 text-white">En preparación</Badge>
                              </div>
                              <div className="text-sm space-y-1">
                                {order.items.map((item, idx) => (
                                  <div key={idx} className="flex justify-between">
                                    <span>
                                      {item.name} x{item.quantity}
                                    </span>
                                    <span>S/ {(item.price * item.quantity).toFixed(2)}</span>
                                  </div>
                                ))}
                                <div className="pt-2 border-t flex justify-between">
                                  <span>Subtotal:</span>
                                  <span>S/ {(order.subtotal || 0).toFixed(2)}</span>
                                </div>
                                <div className="flex justify-between">
                                  <span>Envío:</span>
                                  <span>S/ {(order.envio || 0).toFixed(2)}</span>
                                </div>
                                <div className="font-semibold flex justify-between">
                                  <span>Total:</span>
                                  <span>S/ {(order.total || 0).toFixed(2)}</span>
                                </div>
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    ) : (
                      <p className="text-center text-[#6D4C41] py-8">No hay pedidos en preparación</p>
                    )}
                  </TabsContent>

                  <TabsContent value="rechazado">
                    {getOrdersByStatus("Rechazado").length > 0 ? (
                      <div className="space-y-4">
                        {getOrdersByStatus("Rechazado").map((order) => (
                          <Card key={order.id} className="bg-white">
                            <CardContent className="p-4">
                              <div className="flex justify-between items-start mb-3">
                                <div>
                                  <h4 className="font-bold text-[#3E2723]">Orden #{order.id}</h4>
                                  <p className="text-sm text-[#6D4C41]">{order.userName}</p>
                                  <p className="text-sm text-[#6D4C41]">{order.date}</p>
                                </div>
                                <Badge className="bg-red-500 text-white">Rechazado</Badge>
                              </div>
                              {order.rejectionReason && (
                                <div className="mb-3 p-2 bg-red-50 rounded">
                                  <p className="text-sm text-red-800">
                                    <strong>Motivo:</strong> {order.rejectionReason}
                                  </p>
                                </div>
                              )}
                              <div className="text-sm space-y-1">
                                {order.items.map((item, idx) => (
                                  <div key={idx} className="flex justify-between">
                                    <span>
                                      {item.name} x{item.quantity}
                                    </span>
                                    <span>S/ {(item.price * item.quantity).toFixed(2)}</span>
                                  </div>
                                ))}
                                <div className="pt-2 border-t flex justify-between">
                                  <span>Subtotal:</span>
                                  <span>S/ {(order.subtotal || 0).toFixed(2)}</span>
                                </div>
                                <div className="flex justify-between">
                                  <span>Envío:</span>
                                  <span>S/ {(order.envio || 0).toFixed(2)}</span>
                                </div>
                                <div className="font-semibold flex justify-between">
                                  <span>Total:</span>
                                  <span>S/ {(order.total || 0).toFixed(2)}</span>
                                </div>
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    ) : (
                      <p className="text-center text-[#6D4C41] py-8">No hay pedidos rechazados</p>
                    )}
                  </TabsContent>
                </Tabs>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Products Tab */}
          <TabsContent value="products" className="space-y-4">
            <Card>
              <CardHeader className="flex flex-row items-center justify-between">
                <CardTitle className="text-[#3E2723]">Inventario de Productos</CardTitle>
                <Dialog open={isAddingProduct} onOpenChange={setIsAddingProduct}>
                  <DialogTrigger asChild>
                    <Button className="bg-[#4F5729] hover:bg-[#3E4520] text-white">
                      <Plus className="h-4 w-4 mr-2" />
                      Agregar Producto
                    </Button>
                  </DialogTrigger>
                  <DialogContent className="bg-[#EDE4CC] max-w-md">
                    <DialogHeader>
                      <DialogTitle className="text-[#3E2723]">Nuevo Producto</DialogTitle>
                    </DialogHeader>
                    <div className="space-y-4 py-4">
                      <div className="space-y-2">
                        <Label htmlFor="name" className="text-[#3E2723]">
                          Nombre del Producto
                        </Label>
                        <Input
                          id="name"
                          placeholder="Ej: Mazamorra de Tocosh"
                          value={newProduct.name}
                          onChange={(e) => setNewProduct({ ...newProduct, name: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="description" className="text-[#3E2723]">
                          Descripción
                        </Label>
                        <Textarea
                          id="description"
                          placeholder="Describe el producto..."
                          value={newProduct.description}
                          onChange={(e) => setNewProduct({ ...newProduct, description: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="price" className="text-[#3E2723]">
                          Precio (S/)
                        </Label>
                        <Input
                          id="price"
                          type="number"
                          step="0.01"
                          placeholder="0.00"
                          value={newProduct.price}
                          onChange={(e) => setNewProduct({ ...newProduct, price: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="category" className="text-[#3E2723]">
                          Categoría
                        </Label>
                        <Select
                          value={newProduct.category}
                          onValueChange={(value) =>
                            setNewProduct({ ...newProduct, category: value as Product["category"] })
                          }
                        >
                          <SelectTrigger className="bg-white">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="mazamorra">Mazamorra</SelectItem>
                            <SelectItem value="harina">Harina</SelectItem>
                            <SelectItem value="bebida">Bebida</SelectItem>
                            <SelectItem value="galleta">Galleta</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="imageUpload" className="text-[#3E2723]">
                          Imagen del Producto
                        </Label>
                        <div className="flex items-center gap-2">
                          <Input
                            id="imageUpload"
                            type="file"
                            accept="image/*"
                            onChange={(e) => handleImageUpload(e)}
                            className="bg-white"
                          />
                          <Upload className="h-5 w-5 text-[#6D4C41]" />
                        </div>
                        {newProduct.image && (
                          <div className="mt-2">
                            <img
                              src={newProduct.image || "/placeholder.svg"}
                              alt="Preview"
                              className="w-20 h-20 object-cover rounded-lg"
                            />
                          </div>
                        )}
                      </div>

                      <div className="flex gap-2 pt-4">
                        <Button
                          onClick={handleAddProduct}
                          className="flex-1 bg-[#4F5729] hover:bg-[#3E4520] text-white"
                        >
                          Guardar Producto
                        </Button>
                        <Button onClick={() => setIsAddingProduct(false)} variant="outline" className="flex-1">
                          Cancelar
                        </Button>
                      </div>
                    </div>
                  </DialogContent>
                </Dialog>
              </CardHeader>
              <CardContent>
                <div className="bg-white rounded-lg overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full">
                      <thead>
                        <tr className="border-b border-[#D7CCC8]">
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Imagen</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Nombre</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Precio</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Categoría</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Stock</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#829356]">Acciones</th>
                        </tr>
                      </thead>
                      <tbody>
                        {products.map((product) => (
                          <tr key={product.id} className="border-b border-[#D7CCC8] hover:bg-[#EDE4CC]/30">
                            <td className="py-4 px-6">
                              <div className="w-12 h-12 rounded-full bg-[#DBCFAE] flex items-center justify-center overflow-hidden">
                                <img
                                  src={product.image || "/placeholder.svg"}
                                  alt={product.name}
                                  className="w-full h-full object-cover"
                                />
                              </div>
                            </td>
                            <td className="py-4 px-6 text-[#3E2723]">{product.name}</td>
                            <td className="py-4 px-6 text-[#3E2723]">S/ {(product.price || 0).toFixed(2)}</td>
                            <td className="py-4 px-6 text-[#829356] capitalize">{product.category}</td>
                            <td className="py-4 px-6 text-[#3E2723]">99</td>
                            <td className="py-4 px-6">
                              <Dialog
                                open={editingProduct?.id === product.id}
                                onOpenChange={(open) => {
                                  if (!open) setEditingProduct(null)
                                }}
                              >
                                <DialogTrigger asChild>
                                  <button
                                    onClick={() => setEditingProduct(product)}
                                    className="text-[#829356] hover:text-[#4F5729] font-medium text-sm"
                                  >
                                    Editar
                                  </button>
                                </DialogTrigger>
                                <DialogContent className="bg-[#EDE4CC] max-w-md">
                                  <DialogHeader>
                                    <DialogTitle className="text-[#3E2723]">Editar Producto</DialogTitle>
                                  </DialogHeader>
                                  {editingProduct && (
                                    <div className="space-y-4 py-4">
                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Nombre del Producto</Label>
                                        <Input
                                          value={editingProduct.name}
                                          onChange={(e) =>
                                            setEditingProduct({ ...editingProduct, name: e.target.value })
                                          }
                                          className="bg-white"
                                        />
                                      </div>

                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Descripción</Label>
                                        <Textarea
                                          value={editingProduct.description}
                                          onChange={(e) =>
                                            setEditingProduct({ ...editingProduct, description: e.target.value })
                                          }
                                          className="bg-white"
                                        />
                                      </div>

                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Precio (S/)</Label>
                                        <Input
                                          type="number"
                                          step="0.01"
                                          value={editingProduct.price}
                                          onChange={(e) =>
                                            setEditingProduct({
                                              ...editingProduct,
                                              price: Number.parseFloat(e.target.value),
                                            })
                                          }
                                          className="bg-white"
                                        />
                                      </div>

                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Categoría</Label>
                                        <Select
                                          value={editingProduct.category}
                                          onValueChange={(value) =>
                                            setEditingProduct({
                                              ...editingProduct,
                                              category: value as Product["category"],
                                            })
                                          }
                                        >
                                          <SelectTrigger className="bg-white">
                                            <SelectValue />
                                          </SelectTrigger>
                                          <SelectContent>
                                            <SelectItem value="mazamorra">Mazamorra</SelectItem>
                                            <SelectItem value="harina">Harina</SelectItem>
                                            <SelectItem value="bebida">Bebida</SelectItem>
                                            <SelectItem value="galleta">Galleta</SelectItem>
                                          </SelectContent>
                                        </Select>
                                      </div>

                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Cambiar Imagen</Label>
                                        <div className="flex items-center gap-2">
                                          <Input
                                            type="file"
                                            accept="image/*"
                                            onChange={(e) => handleImageUpload(e, true)}
                                            className="bg-white"
                                          />
                                          <Upload className="h-5 w-5 text-[#6D4C41]" />
                                        </div>
                                        {editingProduct.image && (
                                          <div className="mt-2">
                                            <img
                                              src={editingProduct.image || "/placeholder.svg"}
                                              alt="Preview"
                                              className="w-20 h-20 object-cover rounded-lg"
                                            />
                                          </div>
                                        )}
                                      </div>

                                      <div className="flex gap-2 pt-4">
                                        <Button
                                          onClick={handleUpdateProduct}
                                          className="flex-1 bg-[#4F5729] hover:bg-[#3E4520] text-white"
                                        >
                                          Actualizar Producto
                                        </Button>
                                        <Button
                                          onClick={() => setEditingProduct(null)}
                                          variant="outline"
                                          className="flex-1"
                                        >
                                          Cancelar
                                        </Button>
                                      </div>
                                    </div>
                                  )}
                                </DialogContent>
                              </Dialog>
                              <span className="text-[#829356] mx-1">|</span>
                              <button
                                onClick={() => handleDeleteProduct(product.id)}
                                className="text-[#829356] hover:text-[#4F5729] font-medium text-sm"
                              >
                                Eliminar
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {products.length === 0 && (
                    <div className="text-center py-12">
                      <p className="text-[#6D4C41]">No hay productos registrados aún</p>
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Promotions Tab */}
          <TabsContent value="promotions" className="space-y-4">
            <Card>
              <CardHeader className="flex flex-row items-center justify-between">
                <CardTitle className="text-[#3E2723]">Gestión de Promociones</CardTitle>
                <Dialog open={isAddingPromotion} onOpenChange={setIsAddingPromotion}>
                  <DialogTrigger asChild>
                    <Button className="bg-[#4F5729] hover:bg-[#3E4520] text-white">
                      <Plus className="h-4 w-4 mr-2" />
                      Agregar Promoción
                    </Button>
                  </DialogTrigger>
                  <DialogContent className="bg-[#EDE4CC] max-w-md">
                    <DialogHeader>
                      <DialogTitle className="text-[#3E2723]">Nueva Promoción</DialogTitle>
                    </DialogHeader>
                    <div className="space-y-4 py-4">
                      <div className="space-y-2">
                        <Label htmlFor="promoTitle" className="text-[#3E2723]">
                          Título de la Promoción
                        </Label>
                        <Input
                          id="promoTitle"
                          placeholder="Ej: Combo Familiar"
                          value={newPromotion.title}
                          onChange={(e) => setNewPromotion({ ...newPromotion, title: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="promoDescription" className="text-[#3E2723]">
                          Descripción
                        </Label>
                        <Textarea
                          id="promoDescription"
                          placeholder="Describe la promoción..."
                          value={newPromotion.description}
                          onChange={(e) => setNewPromotion({ ...newPromotion, description: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="grid grid-cols-2 gap-4">
                        <div className="space-y-2">
                          <Label htmlFor="originalPrice" className="text-[#3E2723]">
                            Precio Original (S/)
                          </Label>
                          <Input
                            id="originalPrice"
                            type="number"
                            step="0.01"
                            placeholder="0.00"
                            value={newPromotion.originalPrice}
                            onChange={(e) => setNewPromotion({ ...newPromotion, originalPrice: e.target.value })}
                            className="bg-white"
                          />
                        </div>

                        <div className="space-y-2">
                          <Label htmlFor="discountPrice" className="text-[#3E2723]">
                            Precio con Descuento (S/)
                          </Label>
                          <Input
                            id="discountPrice"
                            type="number"
                            step="0.01"
                            placeholder="0.00"
                            value={newPromotion.discountPrice}
                            onChange={(e) => setNewPromotion({ ...newPromotion, discountPrice: e.target.value })}
                            className="bg-white"
                          />
                        </div>
                      </div>

                      <div className="space-y-2">
                        <Label htmlFor="promoImage" className="text-[#3E2723]">
                          URL de Imagen (opcional)
                        </Label>
                        <Input
                          id="promoImage"
                          placeholder="https://ejemplo.com/imagen.jpg"
                          value={newPromotion.image}
                          onChange={(e) => setNewPromotion({ ...newPromotion, image: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="flex gap-2 pt-4">
                        <Button
                          onClick={handleAddPromotion}
                          className="flex-1 bg-[#4F5729] hover:bg-[#3E4520] text-white"
                        >
                          Guardar Promoción
                        </Button>
                        <Button onClick={() => setIsAddingPromotion(false)} variant="outline" className="flex-1">
                          Cancelar
                        </Button>
                      </div>
                    </div>
                  </DialogContent>
                </Dialog>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {promotions.map((promo) => (
                    <Card key={promo.id} className="overflow-hidden hover:shadow-lg transition-shadow bg-white">
                      <CardContent className="p-0">
                        <div className="relative h-48">
                          <img
                            src={promo.image || "/placeholder.svg"}
                            alt={promo.title}
                            className="w-full h-full object-cover"
                          />
                          <Badge className="absolute top-4 right-4 bg-red-600 text-white text-lg px-3 py-1">
                            -{promo.discount}%
                          </Badge>
                        </div>
                        <div className="p-4">
                          <h3 className="text-lg font-bold text-[#3E2723] mb-2">{promo.title}</h3>
                          <p className="text-sm text-[#6D4C41] mb-3 line-clamp-2">{promo.description}</p>
                          <div className="flex items-center gap-2 mb-4">
                            <span className="text-sm text-[#6D4C41] line-through">
                              S/ {promo.originalPrice.toFixed(2)}
                            </span>
                            <span className="text-xl font-bold text-[#4F5729]">
                              S/ {promo.discountPrice.toFixed(2)}
                            </span>
                          </div>
                          <Button
                            onClick={() => handleDeletePromotion(promo.id)}
                            variant="outline"
                            className="w-full border-red-600 text-red-600 hover:bg-red-600 hover:text-white"
                          >
                            Eliminar Promoción
                          </Button>
                        </div>
                      </CardContent>
                    </Card>
                  ))}
                </div>
                {promotions.length === 0 && (
                  <div className="text-center py-12">
                    <p className="text-[#6D4C41]">No hay promociones registradas aún</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </TabsContent>

          {/* Users Tab */}
          <TabsContent value="users" className="space-y-4">
            <Card>
              <CardHeader className="bg-[#4F5729] text-white rounded-t-lg">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <CardTitle className="text-2xl">Usuarios</CardTitle>
                  <Dialog open={isAddingUser} onOpenChange={setIsAddingUser}>
                    <DialogTrigger asChild>
                      <Button className="bg-white text-[#4F5729] hover:bg-[#EDE4CC]">Agregar Usuario</Button>
                    </DialogTrigger>
                    <DialogContent className="bg-[#EDE4CC] max-w-md">
                      <DialogHeader>
                        <DialogTitle className="text-[#3E2723]">Nuevo Usuario</DialogTitle>
                      </DialogHeader>
                      <div className="space-y-4 py-4">
                        <div className="space-y-2">
                          <Label htmlFor="userName" className="text-[#3E2723]">
                            Nombre Completo
                          </Label>
                          <Input
                            id="userName"
                            placeholder="Ej: Juan Pérez"
                            value={newUser.name}
                            onChange={(e) => setNewUser({ ...newUser, name: e.target.value })}
                            className="bg-white"
                          />
                        </div>

                        <div className="space-y-2">
                          <Label htmlFor="userEmail" className="text-[#3E2723]">
                            Correo Electrónico
                          </Label>
                          <Input
                            id="userEmail"
                            type="email"
                            placeholder="correo@ejemplo.com"
                            value={newUser.email}
                            onChange={(e) => setNewUser({ ...newUser, email: e.target.value })}
                            className="bg-white"
                          />
                        </div>

                        <div className="space-y-2">
                          <Label htmlFor="userPassword" className="text-[#3E2723]">
                            Contraseña
                          </Label>
                          <Input
                            id="userPassword"
                            type="password"
                            placeholder="••••••••"
                            value={newUser.password}
                            onChange={(e) => setNewUser({ ...newUser, password: e.target.value })}
                            className="bg-white"
                          />
                        </div>

                        <div className="space-y-2">
                          <Label htmlFor="userRole" className="text-[#3E2723]">
                            Rol
                          </Label>
                          <Select
                            value={newUser.role}
                            onValueChange={(value) => setNewUser({ ...newUser, role: value as UserRole })}
                          >
                            <SelectTrigger className="bg-white">
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="Cliente">Cliente</SelectItem>
                              <SelectItem value="Vendedor">Vendedor</SelectItem>
                              <SelectItem value="Administrador">Administrador</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>

                        <div className="flex gap-2 pt-4">
                          <Button onClick={handleAddUser} className="flex-1 bg-[#4F5729] hover:bg-[#3E4520] text-white">
                            Crear Usuario
                          </Button>
                          <Button onClick={() => setIsAddingUser(false)} variant="outline" className="flex-1">
                            Cancelar
                          </Button>
                        </div>
                      </div>
                    </DialogContent>
                  </Dialog>
                </div>
              </CardHeader>
              <CardContent className="p-6">
                <div className="mb-6">
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-[#6D4C41] h-5 w-5" />
                    <Input
                      placeholder="Buscar usuarios por nombre o correo electrónico"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="pl-10 bg-white"
                    />
                  </div>
                </div>

                <div className="bg-white rounded-lg overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full">
                      <thead>
                        <tr className="border-b border-[#D7CCC8]">
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Nombre</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Correo Electrónico</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Fecha de Registro</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#3E2723]">Rol</th>
                          <th className="text-left py-4 px-6 text-sm font-medium text-[#829356]">Acciones</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredUsers.map((user) => (
                          <tr key={user.email} className="border-b border-[#D7CCC8] hover:bg-[#EDE4CC]/30">
                            <td className="py-4 px-6 text-[#3E2723]">{user.name}</td>
                            <td className="py-4 px-6 text-[#829356]">{user.email}</td>
                            <td className="py-4 px-6 text-[#3E2723]">{user.registrationDate}</td>
                            <td className="py-4 px-6">
                              <Badge className={`${getRoleBadgeColor(user.role)} rounded-full px-4 py-1`}>
                                {user.role}
                              </Badge>
                            </td>
                            <td className="py-4 px-6">
                              <Dialog>
                                <DialogTrigger asChild>
                                  <button
                                    onClick={() => setEditingUser(user)}
                                    className="text-[#829356] hover:text-[#4F5729] font-medium text-sm"
                                  >
                                    Editar
                                  </button>
                                </DialogTrigger>
                                <DialogContent className="bg-[#EDE4CC] max-w-md">
                                  <DialogHeader>
                                    <DialogTitle className="text-[#3E2723]">Editar Usuario</DialogTitle>
                                  </DialogHeader>
                                  {editingUser && (
                                    <div className="space-y-4 py-4">
                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Nombre Completo</Label>
                                        <Input
                                          value={editingUser.name}
                                          onChange={(e) => setEditingUser({ ...editingUser, name: e.target.value })}
                                          className="bg-white"
                                        />
                                      </div>

                                      <div className="space-y-2">
                                        <Label className="text-[#3E2723]">Rol</Label>
                                        <Select
                                          value={editingUser.role}
                                          onValueChange={(value) =>
                                            setEditingUser({ ...editingUser, role: value as UserRole })
                                          }
                                        >
                                          <SelectTrigger className="bg-white">
                                            <SelectValue />
                                          </SelectTrigger>
                                          <SelectContent>
                                            <SelectItem value="Cliente">Cliente</SelectItem>
                                            <SelectItem value="Vendedor">Vendedor</SelectItem>
                                            <SelectItem value="Administrador">Administrador</SelectItem>
                                          </SelectContent>
                                        </Select>
                                      </div>

                                      <div className="flex gap-2 pt-4">
                                        <Button
                                          onClick={handleUpdateUser}
                                          className="flex-1 bg-[#4F5729] hover:bg-[#3E4520] text-white"
                                        >
                                          Guardar Cambios
                                        </Button>
                                        <Button
                                          onClick={() => setEditingUser(null)}
                                          variant="outline"
                                          className="flex-1"
                                        >
                                          Cancelar
                                        </Button>
                                      </div>
                                    </div>
                                  )}
                                </DialogContent>
                              </Dialog>
                              <span className="text-[#829356] mx-1">|</span>
                              <button
                                onClick={() => handleToggleUserStatus(user.email)}
                                className="text-[#829356] hover:text-[#4F5729] font-medium text-sm"
                              >
                                {user.isActive ? "Activo" : "Inactivo"}
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {filteredUsers.length === 0 && (
                    <div className="text-center py-12">
                      <p className="text-[#6D4C41]">
                        {searchQuery ? "No se encontraron usuarios" : "No hay usuarios registrados aún"}
                      </p>
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="reports" className="space-y-4">
            <Card>
              <CardHeader className="bg-[#4F5729] text-white rounded-t-lg">
                <CardTitle className="text-2xl">Generar Reportes de Pedidos</CardTitle>
              </CardHeader>
              <CardContent className="p-6">
                <div className="space-y-6">
                  <div>
                    <h3 className="text-lg font-semibold text-[#3E2723] mb-4">Filtrar por</h3>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                      <div className="space-y-2">
                        <Label className="text-[#3E2723]">Fecha de Inicio</Label>
                        <Input
                          type="date"
                          value={reportFilters.startDate}
                          onChange={(e) => setReportFilters({ ...reportFilters, startDate: e.target.value })}
                          className="bg-white"
                        />
                      </div>

                      <div className="space-y-2">
                        <Label className="text-[#3E2723]">Fecha de Fin</Label>
                        <Input
                          type="date"
                          value={reportFilters.endDate}
                          onChange={(e) => setReportFilters({ ...reportFilters, endDate: e.target.value })}
                          className="bg-white"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                      <div className="space-y-2">
                        <Label className="text-[#3E2723]">Estado del Pedido</Label>
                        <Select
                          value={reportFilters.status}
                          onValueChange={(value) =>
                            setReportFilters({ ...reportFilters, status: value as Order["status"] | "Todos" })
                          }
                        >
                          <SelectTrigger className="bg-white">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Todos">Todos</SelectItem>
                            <SelectItem value="Completado">Completado</SelectItem>
                            <SelectItem value="Pendiente">Pendiente</SelectItem>
                            <SelectItem value="En preparación">En preparación</SelectItem>
                            <SelectItem value="Listo para envío">Listo para envío</SelectItem>
                            <SelectItem value="Rechazado">Rechazado</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>

                      <div className="space-y-2">
                        <Label className="text-[#3E2723]">Cliente</Label>
                        <Select
                          value={reportFilters.customer}
                          onValueChange={(value) => setReportFilters({ ...reportFilters, customer: value })}
                        >
                          <SelectTrigger className="bg-white">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Todos">Todos</SelectItem>
                            {getUniqueCustomers().map((customer) => (
                              <SelectItem key={customer} value={customer}>
                                {customer}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>
                    </div>

                    <div className="space-y-2 mb-4">
                      <Label className="text-[#3E2723]">Tipo de Producto</Label>
                      <Input
                        placeholder="Seleccionar producto"
                        value={reportFilters.productType}
                        onChange={(e) => setReportFilters({ ...reportFilters, productType: e.target.value })}
                        className="bg-white"
                      />
                    </div>

                    <div className="mb-6">
                      <h3 className="text-lg font-semibold text-[#3E2723] mb-4">O Ingrese Número de Pedido</h3>
                      <div className="space-y-2">
                        <Label className="text-[#3E2723]">Número de Pedido</Label>
                        <Input
                          placeholder="Ingrese el número de pedido"
                          value={orderNumberSearch}
                          onChange={(e) => setOrderNumberSearch(e.target.value)}
                          className="bg-white"
                        />
                      </div>
                    </div>

                    <Button onClick={handleGenerateReport} className="bg-[#4F5729] hover:bg-[#3E4520] text-white">
                      Generar Reporte
                    </Button>
                  </div>

                  {reportResults.length > 0 && (
                    <div className="space-y-4">
                      <h3 className="text-lg font-semibold text-[#3E2723]">Reporte Generado</h3>

                      <div className="bg-white rounded-lg overflow-hidden">
                        <div className="overflow-x-auto">
                          <table className="w-full">
                            <thead className="bg-[#4F5729] text-white">
                              <tr>
                                <th className="text-left py-3 px-4 text-sm font-medium">Número de Pedido</th>
                                <th className="text-left py-3 px-4 text-sm font-medium">Cliente</th>
                                <th className="text-left py-3 px-4 text-sm font-medium">Fecha</th>
                                <th className="text-left py-3 px-4 text-sm font-medium">Total Pagado</th>
                                <th className="text-left py-3 px-4 text-sm font-medium">Estado</th>
                                <th className="text-left py-3 px-4 text-sm font-medium">Medio de Pago</th>
                              </tr>
                            </thead>
                            <tbody>
                              {reportResults.map((order) => (
                                <tr key={order.id} className="border-b border-[#D7CCC8]">
                                  <td className="py-3 px-4 text-[#3E2723]">{order.id}</td>
                                  <td className="py-3 px-4 text-[#829356]">{order.userName}</td>
                                  <td className="py-3 px-4 text-[#829356]">{order.date}</td>
                                  <td className="py-3 px-4 text-[#3E2723]">S/ {order.total.toFixed(2)}</td>
                                  <td className="py-3 px-4">
                                    <span className="text-[#829356]">{order.status}</span>
                                  </td>
                                  <td className="py-3 px-4 text-[#3E2723]">
                                    {order.paymentMethod === "card" ? "Tarjeta" : "Yape"}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>

                      <div className="flex gap-3">
                        <Button
                          onClick={() => exportToPDF(reportResults)}
                          className="bg-red-600 hover:bg-red-700 text-white"
                        >
                          <FileDown className="h-4 w-4 mr-2" />
                          Descargar PDF
                        </Button>
                        <Button
                          onClick={() => exportToExcel(reportResults)}
                          className="bg-green-600 hover:bg-green-700 text-white"
                        >
                          <FileDown className="h-4 w-4 mr-2" />
                          Descargar Excel
                        </Button>
                      </div>
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>
      </main>
    </div>
  )
}
