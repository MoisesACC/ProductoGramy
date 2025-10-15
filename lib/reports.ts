import type { Order } from "./orders"
import { getOrders } from "./orders"
import jsPDF from "jspdf"
import "jspdf-autotable"
import * as XLSX from "xlsx"

export interface ReportFilters {
  startDate?: string
  endDate?: string
  status?: Order["status"] | "Todos"
  customer?: string
  productType?: string
}

export function generateOrderReport(filters: ReportFilters): Order[] {
  let orders = getOrders()

  // Filter by date range
  if (filters.startDate) {
    orders = orders.filter((order) => {
      const orderDate = new Date(order.date.split("/").reverse().join("-"))
      const startDate = new Date(filters.startDate!)
      return orderDate >= startDate
    })
  }

  if (filters.endDate) {
    orders = orders.filter((order) => {
      const orderDate = new Date(order.date.split("/").reverse().join("-"))
      const endDate = new Date(filters.endDate!)
      return orderDate <= endDate
    })
  }

  // Filter by status
  if (filters.status && filters.status !== "Todos") {
    orders = orders.filter((order) => order.status === filters.status)
  }

  // Filter by customer
  if (filters.customer && filters.customer !== "Todos") {
    orders = orders.filter((order) => order.userName === filters.customer || order.userEmail === filters.customer)
  }

  // Filter by product type
  if (filters.productType && filters.productType !== "Todos") {
    orders = orders.filter((order) =>
      order.items.some((item) => item.name.toLowerCase().includes(filters.productType!.toLowerCase())),
    )
  }

  return orders
}

export function getUniqueCustomers(): string[] {
  const orders = getOrders()
  const customers = new Set<string>()
  orders.forEach((order) => {
    customers.add(order.userName)
  })
  return Array.from(customers)
}

export function exportToPDF(orders: Order[]): void {
  const doc = new jsPDF()

  // Add company logo and header
  doc.setFontSize(20)
  doc.setTextColor(79, 87, 41) // #4F5729
  doc.text("GRAMY", 14, 20)

  doc.setFontSize(16)
  doc.text("Reporte de Pedidos", 14, 30)

  doc.setFontSize(10)
  doc.setTextColor(0, 0, 0)
  doc.text(`Fecha de generación: ${new Date().toLocaleDateString("es-PE")}`, 14, 38)
  doc.text(`Total de pedidos: ${orders.length}`, 14, 44)

  // Calculate total revenue
  const totalRevenue = orders.reduce((sum, order) => sum + (order.total || 0), 0)
  doc.text(`Ingresos totales: S/ ${totalRevenue.toFixed(2)}`, 14, 50)

  // Add table with order details
  const tableData = orders.map((order) => [
    order.id,
    order.userName,
    order.date,
    `S/ ${(order.subtotal || 0).toFixed(2)}`,
    `S/ ${(order.envio || 0).toFixed(2)}`,
    `S/ ${(order.total || 0).toFixed(2)}`,
    order.status,
    order.paymentMethod === "card" ? "Tarjeta" : "Yape",
  ])
  ;(doc as any).autoTable({
    startY: 58,
    head: [["N° Pedido", "Cliente", "Fecha", "Subtotal", "Envío", "Total", "Estado", "Pago"]],
    body: tableData,
    theme: "grid",
    headStyles: {
      fillColor: [79, 87, 41], // #4F5729
      textColor: [255, 255, 255],
      fontSize: 9,
      fontStyle: "bold",
    },
    bodyStyles: {
      fontSize: 8,
    },
    alternateRowStyles: {
      fillColor: [237, 228, 204], // #EDE4CC
    },
    margin: { top: 58 },
  })

  // Add footer
  const pageCount = (doc as any).internal.getNumberOfPages()
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i)
    doc.setFontSize(8)
    doc.setTextColor(109, 76, 65)
    doc.text(
      `Página ${i} de ${pageCount}`,
      doc.internal.pageSize.getWidth() / 2,
      doc.internal.pageSize.getHeight() - 10,
      { align: "center" },
    )
  }

  // Download the PDF
  doc.save(`reporte_pedidos_${new Date().toISOString().split("T")[0]}.pdf`)
}

export function exportToExcel(orders: Order[]): void {
  const wb = XLSX.utils.book_new()

  // Create company header data
  const headerData = [
    ["GRAMY - Reporte de Pedidos"],
    [`Fecha de generación: ${new Date().toLocaleDateString("es-PE")}`],
    [`Total de pedidos: ${orders.length}`],
    [`Ingresos totales: S/ ${orders.reduce((sum, order) => sum + (order.total || 0), 0).toFixed(2)}`],
    [], // Empty row
    ["Número de Pedido", "Cliente", "Email", "Fecha", "Subtotal", "Envío", "Total", "Estado", "Método de Pago"],
  ]

  // Create order data rows
  const orderData = orders.map((order) => [
    order.id,
    order.userName,
    order.userEmail,
    order.date,
    (order.subtotal || 0).toFixed(2),
    (order.envio || 0).toFixed(2),
    (order.total || 0).toFixed(2),
    order.status,
    order.paymentMethod === "card" ? "Tarjeta" : "Yape",
  ])

  // Combine header and data
  const wsData = [...headerData, ...orderData]

  // Create worksheet
  const ws = XLSX.utils.aoa_to_sheet(wsData)

  // Set column widths
  ws["!cols"] = [
    { wch: 15 }, // Número de Pedido
    { wch: 20 }, // Cliente
    { wch: 25 }, // Email
    { wch: 12 }, // Fecha
    { wch: 10 }, // Subtotal
    { wch: 10 }, // Envío
    { wch: 10 }, // Total
    { wch: 15 }, // Estado
    { wch: 15 }, // Método de Pago
  ]

  // Style the header rows
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1")
  for (let R = 0; R <= 5; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C })
      if (!ws[cellAddress]) continue

      if (R === 0) {
        // Title row
        ws[cellAddress].s = {
          font: { bold: true, sz: 16, color: { rgb: "4F5729" } },
          alignment: { horizontal: "center" },
        }
      } else if (R === 5) {
        // Column headers
        ws[cellAddress].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "4F5729" } },
          alignment: { horizontal: "center" },
        }
      }
    }
  }

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, "Reporte de Pedidos")

  // Generate Excel file and download
  XLSX.writeFile(wb, `reporte_pedidos_${new Date().toISOString().split("T")[0]}.xlsx`)
}

function generateCSV(orders: Order[]): string {
  const headers = ["Número de Pedido", "Cliente", "Fecha", "Total Pagado", "Estado", "Medio de Pago"]
  const rows = orders.map((order) => [
    order.id,
    order.userName,
    order.date,
    `S/ ${order.total.toFixed(2)}`,
    order.status,
    order.paymentMethod === "card" ? "Tarjeta" : "Yape",
  ])

  const csvRows = [headers, ...rows].map((row) => row.join(",")).join("\n")
  return csvRows
}
