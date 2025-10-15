export interface User {
  email: string
  name: string
  password: string
  isAdmin: boolean
  role: UserRole
  registrationDate: string
  isActive: boolean
}

export type UserRole = "Administrador" | "Vendedor" | "Cliente"

// Initialize with admin account
const ADMIN_USER: User = {
  email: "melanieAdmin@gmail.com",
  name: "Melanie Admin",
  password: "12345",
  isAdmin: true,
  role: "Administrador",
  registrationDate: new Date().toLocaleDateString("es-PE"),
  isActive: true,
}

export function initializeAuth() {
  if (typeof window === "undefined") return

  const users = getUsers()
  // Add admin if not exists
  if (!users.find((u) => u.email === ADMIN_USER.email)) {
    users.push(ADMIN_USER)
    localStorage.setItem("gramy_users", JSON.stringify(users))
  }
}

export function getUsers(): User[] {
  if (typeof window === "undefined") return []
  const users = localStorage.getItem("gramy_users")
  return users ? JSON.parse(users) : []
}

export function registerUser(email: string, name: string, password: string): boolean {
  const users = getUsers()

  // Check if user already exists
  if (users.find((u) => u.email === email)) {
    return false
  }

  users.push({
    email,
    name,
    password,
    isAdmin: false,
    role: "Cliente",
    registrationDate: new Date().toLocaleDateString("es-PE"),
    isActive: true,
  })
  localStorage.setItem("gramy_users", JSON.stringify(users))
  return true
}

export function loginUser(email: string, password: string): User | null {
  const users = getUsers()
  const user = users.find((u) => u.email === email && u.password === password)

  if (user) {
    // Store current user session
    const { password: _, ...userWithoutPassword } = user
    localStorage.setItem("gramy_current_user", JSON.stringify(userWithoutPassword))
    return user
  }

  return null
}

export function getCurrentUser(): Omit<User, "password"> | null {
  if (typeof window === "undefined") return null
  const user = localStorage.getItem("gramy_current_user")
  return user ? JSON.parse(user) : null
}

export function logoutUser() {
  if (typeof window === "undefined") return
  localStorage.removeItem("gramy_current_user")
}

export function isAuthenticated(): boolean {
  return getCurrentUser() !== null
}

export function updateUser(email: string, updates: Partial<User>): boolean {
  const users = getUsers()
  const userIndex = users.findIndex((u) => u.email === email)

  if (userIndex === -1) return false

  users[userIndex] = { ...users[userIndex], ...updates }
  localStorage.setItem("gramy_users", JSON.stringify(users))
  return true
}

export function createUser(email: string, name: string, password: string, role: UserRole, isAdmin = false): boolean {
  const users = getUsers()

  if (users.find((u) => u.email === email)) {
    return false
  }

  users.push({
    email,
    name,
    password,
    isAdmin,
    role,
    registrationDate: new Date().toLocaleDateString("es-PE"),
    isActive: true,
  })
  localStorage.setItem("gramy_users", JSON.stringify(users))
  return true
}

export function toggleUserStatus(email: string): boolean {
  const users = getUsers()
  const userIndex = users.findIndex((u) => u.email === email)

  if (userIndex === -1) return false

  users[userIndex].isActive = !users[userIndex].isActive
  localStorage.setItem("gramy_users", JSON.stringify(users))
  return true
}
