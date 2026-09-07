import Link from "next/link";
import { Card, CardHeader, CardTitle, CardDescription, CardContent } from "@/components/ui/card";
import { ArrowRight, FileSpreadsheet, Scale } from "lucide-react";

const menus = [
  {
    title: "แปลงเอกสาร",
    englishTitle: "Document Converter",
    description: "แปลงเอกสารภาษี, ประกันสังคม และ SCB Bank Statement (PDF) เป็น CSV / XLSX",
    href: "/convert",
    icon: FileSpreadsheet,
    badge: "เครื่องมือ",
  },
  {
    title: "กระทบยอด",
    englishTitle: "Reconciliation",
    description: "กระทบยอดภาษีและประกันสังคม 12 เดือน (ภ.ง.ด. 1 vs สปส. 1-10) เป็น CSV / XLSX",
    href: "/reconcile",
    icon: Scale,
    badge: "ตรวจสอบภาษี",
  },
];

export default function HomePage() {
  return (
    <main className="min-h-screen bg-background text-foreground flex flex-col justify-center items-center px-4 py-12 sm:px-6 lg:px-8">
      <div className="w-full max-w-4xl space-y-10">
        <header className="text-center space-y-3">
          <h1 className="text-4xl font-extrabold tracking-tight sm:text-5xl">
            ระบบบัญชี
          </h1>
          <p className="text-muted-foreground text-base sm:text-lg">
            กรุณาเลือกสาขาหรือบริการที่ต้องการเข้าใช้งาน
          </p>
        </header>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 max-w-3xl mx-auto">
          {menus.map((menu) => {
            const Icon = menu.icon;
            return (
              <Link
                key={menu.title}
                href={menu.href}
                className="group block focus:outline-none focus-visible:ring-2 focus-visible:ring-primary rounded-xl"
              >
                <Card className="h-full transition-all duration-200 border-border hover:border-primary/50 hover:shadow-lg hover:-translate-y-1 cursor-pointer bg-card">
                  <CardHeader className="space-y-3">
                    <div className="flex items-center justify-between">
                      <div className="p-3 rounded-lg bg-primary/10 text-primary group-hover:bg-primary group-hover:text-primary-foreground transition-colors duration-200">
                        <Icon className="size-6" />
                      </div>
                      <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-muted text-muted-foreground">
                        {menu.badge}
                      </span>
                    </div>
                    <div>
                      <CardTitle className="text-2xl font-bold text-card-foreground group-hover:text-primary transition-colors">
                        {menu.title}
                      </CardTitle>
                      <span className="text-xs font-medium text-muted-foreground">
                        {menu.englishTitle}
                      </span>
                    </div>
                  </CardHeader>
                  <CardContent className="flex items-center justify-between pt-2">
                    <CardDescription className="text-sm">
                      {menu.description}
                    </CardDescription>
                    <div className="flex items-center text-primary text-sm font-medium opacity-0 group-hover:opacity-100 transition-all duration-200 -translate-x-2 group-hover:translate-x-0 shrink-0 ml-2">
                      <span>เข้าสู่ระบบ</span>
                      <ArrowRight className="size-4 ml-1" />
                    </div>
                  </CardContent>
                </Card>
              </Link>
            );
          })}
        </div>
      </div>
    </main>
  );
}