import type { CartItem } from "@/components/cart-provider"

export interface Order {
  id: string
  userId: string
  userEmail: string
  userName: string
  date: string
  status: "Pendiente" | "En preparación" | "Listo para envío" | "Completado" | "Rechazado"
  total: number
  subtotal: number
  envio: number
  items: CartItem[]
  paymentMethod: "card" | "yape"
  rejectionReason?: string
}

export function createOrder(
  userId: string,
  userEmail: string,
  userName: string,
  items: CartItem[],
  total: number,
  paymentMethod: "card" | "yape",
): Order {
  const subtotal = items.reduce((sum, item) => sum + item.price * item.quantity, 0)
  const envio = 5.0

  const order: Order = {
    id: Math.floor(10000 + Math.random() * 90000).toString(),
    userId,
    userEmail,
    userName,
    date: new Date().toLocaleDateString("es-PE"),
    status: "Pendiente",
    total,
    subtotal,
    envio,
    items,
    paymentMethod,
  }

  const orders = getOrders()
  orders.push(order)
  localStorage.setItem("gramy_orders", JSON.stringify(orders))

  return order
}

export function getOrders(): Order[] {
  if (typeof window === "undefined") return []
  const orders = localStorage.getItem("gramy_orders")
  return orders ? JSON.parse(orders) : []
}

export function getUserOrders(userId: string): Order[] {
  return getOrders().filter((order) => order.userId === userId)
}

export function updateOrderStatus(orderId: string, status: Order["status"], rejectionReason?: string): boolean {
  const orders = getOrders()
  const orderIndex = orders.findIndex((o) => o.id === orderId)

  if (orderIndex === -1) return false

  orders[orderIndex].status = status
  if (rejectionReason) {
    orders[orderIndex].rejectionReason = rejectionReason
  }

  localStorage.setItem("gramy_orders", JSON.stringify(orders))
  return true
}

export function getTotalRevenue(): number {
  const orders = getOrders()
  return orders.filter((order) => order.status === "Completado").reduce((sum, order) => sum + (order.total || 0), 0)
}

export function getPendingOrdersCount(): number {
  return getOrders().filter((order) => order.status === "Pendiente").length
}
