import Link from "next/link"
import { Card, CardContent } from "@/components/ui/card"
import Image from "next/image"

const categories = [
 {
    name: "Mazamorra de Tocosh",
    image: "/mazamorra-tocosh-dessert.jpg",
    href: "/productos?category=mazamorra",
  },
  {
    name: "Harina de Tocosh",
    image: "/tocosh-flour-package.jpg",
    href: "/productos?category=harina",
  },
  {
    name: "Bebida de Tocosh",
    image: "/bebidaTocosh.png",
    href: "/productos?category=bebida",
  },
  {
    name: "Galleta de Tocosh",
    image: "/galletaTocosh.png",
    href: "/productos?category=galleta",
  },
]

export function ProductCategories() {
  return (
    <section>
      <h2 className="text-2xl md:text-3xl font-bold text-[#3E2723] mb-4 md:mb-6">Categorías</h2>
      <div className="grid grid-cols-2 sm:grid-cols-2 md:grid-cols-4 gap-3 md:gap-4 lg:gap-6">
        {categories.map((category) => (
          <Link key={category.name} href={category.href}>
            <Card className="overflow-hidden hover:shadow-lg transition-shadow cursor-pointer">
              <CardContent className="p-0">
                <div className="relative h-32 sm:h-36 md:h-40">
                  <Image src={category.image || "/placeholder.svg"} alt={category.name} fill className="object-cover" />
                </div>
                <div className="p-3 md:p-4 bg-white">
                  <h3 className="font-semibold text-[#3E2723] text-center text-sm md:text-base">{category.name}</h3>
                </div>
              </CardContent>
            </Card>
          </Link>
        ))}
      </div>
    </section>
  )
}
