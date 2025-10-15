import { Card, CardContent } from "@/components/ui/card"
import { Heart, Shield, Sparkles } from "lucide-react"

export function TocoshBenefits() {
  return (
    <section className="py-12">
      <h2 className="text-3xl font-bold text-[#3E2723] mb-6 text-center">
        Explora más sobre los beneficios del Tocosh
      </h2>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <Card className="bg-white">
          <CardContent className="p-6 text-center">
            <div className="flex justify-center mb-4">
              <div className="w-16 h-16 rounded-full bg-[#DBCFAE] flex items-center justify-center">
                <Heart className="h-8 w-8 text-[#4F5729]" />
              </div>
            </div>
            <h3 className="font-semibold text-[#3E2723] mb-2">Salud Digestiva</h3>
            <p className="text-sm text-[#6D4C41]">
              Rico en probióticos naturales que mejoran la flora intestinal y la digestión
            </p>
          </CardContent>
        </Card>

        <Card className="bg-white">
          <CardContent className="p-6 text-center">
            <div className="flex justify-center mb-4">
              <div className="w-16 h-16 rounded-full bg-[#DBCFAE] flex items-center justify-center">
                <Shield className="h-8 w-8 text-[#4F5729]" />
              </div>
            </div>
            <h3 className="font-semibold text-[#3E2723] mb-2">Sistema Inmune</h3>
            <p className="text-sm text-[#6D4C41]">
              Fortalece las defensas naturales del cuerpo con sus propiedades antibióticas
            </p>
          </CardContent>
        </Card>

        <Card className="bg-white">
          <CardContent className="p-6 text-center">
            <div className="flex justify-center mb-4">
              <div className="w-16 h-16 rounded-full bg-[#DBCFAE] flex items-center justify-center">
                <Sparkles className="h-8 w-8 text-[#4F5729]" />
              </div>
            </div>
            <h3 className="font-semibold text-[#3E2723] mb-2">Tradición Andina</h3>
            <p className="text-sm text-[#6D4C41]">Alimento ancestral usado por generaciones en los Andes peruanos</p>
          </CardContent>
        </Card>
      </div>
    </section>
  )
}
