export type UserRole = "Administrador" | "Vendedor" | "Cliente"

export interface User {
  email: string
  name: string
  role: UserRole
  isAdmin: boolean
  isActive: boolean
  registrationDate: string
}

export interface Order {
  id: string
  userId: string
  userName: string
  userEmail: string
  items: Array<{
    id: string
    name: string
    price: number
    quantity: number
    image: string
  }>
  total: number
  status: "Pendiente" | "En preparación" | "Listo para envío" | "Completado" | "Rechazado"
  date: string
  paymentMethod: "card" | "yape"
  rejectionReason?: string
}

const USERS_STORAGE_KEY = "gramy_users"

export function getUsers(): User[] {
  if (typeof window === "undefined") return []

  const stored = localStorage.getItem(USERS_STORAGE_KEY)
  if (!stored) return []

  try {
    return JSON.parse(stored)
  } catch {
    return []
  }
}

export function createUser(email: string, name: string, password: string, role: UserRole, isAdmin: boolean): boolean {
  // Check if user already exists
  const existingUsers = JSON.parse(localStorage.getItem("gramy_users_auth") || "[]")
  if (existingUsers.some((u: any) => u.email === email)) {
    return false
  }

  // Add to auth storage
  const newAuthUser = {
    email,
    name,
    password,
    isAdmin,
  }
  existingUsers.push(newAuthUser)
  localStorage.setItem("gramy_users_auth", JSON.stringify(existingUsers))

  // Add to users list
  const users = getUsers()
  const newUser: User = {
    email,
    name,
    role,
    isAdmin,
    isActive: true,
    registrationDate: new Date().toLocaleDateString("es-PE"),
  }
  users.push(newUser)
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users))

  return true
}

export function updateUser(email: string, updates: Partial<User>): void {
  const users = getUsers()
  const index = users.findIndex((u) => u.email === email)

  if (index !== -1) {
    users[index] = { ...users[index], ...updates }
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users))

    // Also update auth storage if admin status changed
    if (updates.isAdmin !== undefined) {
      const authUsers = JSON.parse(localStorage.getItem("gramy_users_auth") || "[]")
      const authIndex = authUsers.findIndex((u: any) => u.email === email)
      if (authIndex !== -1) {
        authUsers[authIndex].isAdmin = updates.isAdmin
        localStorage.setItem("gramy_users_auth", JSON.stringify(authUsers))
      }
    }
  }
}

export function toggleUserStatus(email: string): void {
  const users = getUsers()
  const index = users.findIndex((u) => u.email === email)

  if (index !== -1) {
    users[index].isActive = !users[index].isActive
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users))
  }
}

// Sync users from auth storage to users list
export function syncUsersFromAuth(): void {
  const authUsers = JSON.parse(localStorage.getItem("gramy_users_auth") || "[]")
  const existingUsers = getUsers()

  authUsers.forEach((authUser: any) => {
    const exists = existingUsers.some((u) => u.email === authUser.email)
    if (!exists) {
      const newUser: User = {
        email: authUser.email,
        name: authUser.name,
        role: authUser.isAdmin ? "Administrador" : "Cliente",
        isAdmin: authUser.isAdmin,
        isActive: true,
        registrationDate: new Date().toLocaleDateString("es-PE"),
      }
      existingUsers.push(newUser)
    }
  })

  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(existingUsers))
}
