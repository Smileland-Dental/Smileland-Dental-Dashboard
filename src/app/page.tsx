import { GalleryVerticalEnd } from "lucide-react"
import Image from 'next/image';
import { LoginForm } from "@/components/login-form"
import background from '../../public/background.jpg';
import dashboard from '../../public/dashboard.png';

export default function LoginPage() {
  return (
    <div className="grid min-h-svh lg:grid-cols-2">
      <div className="flex flex-col gap-4 p-6 md:p-10">
        <div className="flex justify-center items-center gap-2 md:justify-start">
            <div className="bg-primary text-primary-foreground flex size-9 items-center justify-center rounded-md">
              <Image src={dashboard} className="size-8" alt="Icon for Smileland" />
            </div>
            <p className="font-bold underline">Smileland Dental</p>
        </div>
        {/*<div className="flex items-center justify-center gap-2">
          <div className="bg-primary text-primary-foreground flex size-9 items-center justify-center rounded-md">
            <Image src={dashboard} className="size-8" alt="Icon for Smileland" />
          </div>
          <p className="font-bold underline">Smileland Dental</p>
        </div>*/}
        <div className="flex flex-1 items-center justify-center">
          <div className="w-full max-w-xs">
            <LoginForm />
          </div>
        </div>
      </div>
      <div className="bg-muted relative hidden lg:block">
          <Image 
            src={background}
            className="absolute inset-0 h-full w-full object-cover dark:brightness-[0.2] dark:grayscale"
            alt="Screenshots of the dashboard project showing mobile version"
          />
      </div>
    </div>
  )
}
