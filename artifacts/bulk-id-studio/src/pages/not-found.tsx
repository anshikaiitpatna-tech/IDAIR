import { AlertCircle } from 'lucide-react';

export default function NotFound() {
  return (
    <div className="min-h-screen w-full flex items-center justify-center p-6"
         style={{
           backgroundImage:
             'linear-gradient(180deg, #7ec8f5 0%, #b8e0f8 40%, #e8f4fc 70%, #f0e6c8 100%)',
         }}>
      <div className="pixel-panel max-w-md w-full mx-4 p-6 text-center">
        <div className="flex justify-center mb-4">
          <div className="grid h-14 w-14 place-items-center pixel-border-sm bg-[#ff6b6b] text-white">
            <AlertCircle size={28} strokeWidth={2.5} />
          </div>
        </div>

        <h1 className="pixel-font text-[14px] text-[#1a1a2e] mb-3 leading-relaxed">
          404<br />PAGE NOT FOUND
        </h1>

        <p className="text-sm text-[#1a1a2e]/70 mb-6">
          This screen doesn’t exist in the game world.
        </p>

        <a
          href="/"
          className="pixel-btn inline-block bg-[#ffe566] px-6 py-2.5 text-[12px] font-bold text-[#1a1a2e] hover:bg-[#ffd700]"
        >
          BACK TO START
        </a>
      </div>
    </div>
  );
}
