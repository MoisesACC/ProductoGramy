import { Navbar } from "@/components/navbar"
import { HeroCarousel } from "@/components/hero-carousel"
import { ProductCategories } from "@/components/product-categories"
import { PopularProducts } from "@/components/popular-products"
import { TocoshBenefits } from "@/components/tocosh-benefits"

export default function HomePage() {
  return (
    <div className="min-h-screen bg-[#EDE4CC]">
      <Navbar />
      <main className="container mx-auto px-4 py-6 md:py-8 lg:py-12 space-y-8 md:space-y-12 max-w-7xl">
        <HeroCarousel />
        <ProductCategories />
        <PopularProducts />
        <TocoshBenefits />
      </main>
    </div>
  )
}
