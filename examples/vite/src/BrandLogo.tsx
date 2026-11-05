import { useId } from 'react';

// "PPTX Editor by EigenPal" lockup. DOCX Editor's icon is a document page
// (docx-editor/examples/shared/BrandLogo.tsx); this one is a presentation slide in the same
// family style: the same outline, paper, and label badge, with a 16:9 frame, a title, bullet
// lines, and a bar chart, and a `.PPTX` label in orange (#C43E1C) instead of DOCX Editor's
// blue (#185ABD). The EigenPal mark and spacing are unchanged.

function PptxIcon() {
  return (
    <svg
      viewBox="0 0 100 84"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden="true"
      style={{ width: '30px', height: 'auto', flexShrink: 0 }}
    >
      <path
        d="M2 2H98V60H2V2Z"
        fill="#FAF9F8"
        stroke="#201F1E"
        strokeWidth="4"
        strokeLinejoin="round"
      />
      <rect x="12" y="11" width="42" height="6" fill="#201F1E" opacity="0.25" />
      <rect x="12" y="26" width="4" height="4" fill="#201F1E" opacity="0.25" />
      <rect x="20" y="26" width="26" height="4" fill="#201F1E" opacity="0.15" />
      <rect x="12" y="36" width="4" height="4" fill="#201F1E" opacity="0.25" />
      <rect x="20" y="36" width="22" height="4" fill="#201F1E" opacity="0.15" />
      <rect x="12" y="46" width="4" height="4" fill="#201F1E" opacity="0.25" />
      <rect x="20" y="46" width="24" height="4" fill="#201F1E" opacity="0.15" />
      <rect x="56" y="22" width="32" height="30" fill="#FBE3D6" />
      <rect x="61" y="38" width="6" height="10" fill="#C43E1C" opacity="0.6" />
      <rect x="70" y="32" width="6" height="16" fill="#C43E1C" opacity="0.8" />
      <rect x="79" y="26" width="6" height="22" fill="#C43E1C" />
      <rect x="28" y="66" width="44" height="16" fill="#C43E1C" stroke="#201F1E" strokeWidth="3" />
      <text
        x="50"
        y="78"
        textAnchor="middle"
        fill="#FAF9F8"
        fontSize="11"
        fontWeight="900"
        fontFamily="system-ui, sans-serif"
      >
        .PPTX
      </text>
    </svg>
  );
}

export function EigenpalLogo() {
  const uid = useId();
  const [g0, g1, g2] = [`${uid}-g0`, `${uid}-g1`, `${uid}-g2`];
  return (
    <svg
      viewBox="0 0 69 74"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      style={{ width: '10px', height: '11px', flexShrink: 0 }}
      aria-hidden="true"
    >
      <path
        d="M32.364 45.8825C35.29 47.5714 37.0926 50.6931 37.0926 54.0715V68.5491C37.0926 71.1596 34.9756 73.276 32.3651 73.2766C29.7542 73.2766 27.6376 71.16 27.6376 68.5491V46.5581C27.6376 45.7883 26.8043 45.3072 26.1376 45.6921L7.09088 56.6887C4.83001 57.9931 1.9397 57.218 0.6344 54.9574C-0.670978 52.6965 0.10259 49.8042 2.36336 48.4986L14.8956 41.2595C17.8209 39.5698 21.4254 39.5692 24.3513 41.2579L32.364 45.8825Z"
        fill={`url(#${g0})`}
      />
      <path
        d="M0.6344 18.3191C1.93969 16.0585 4.82996 15.2833 7.09088 16.5879L27.6376 28.4482V28.4436L41.8202 36.6337L41.8178 36.636L62.3623 48.4986C64.6233 49.8041 65.3989 52.6963 64.0935 54.9574C63.0062 56.8407 60.8194 57.6884 58.812 57.1665L1.35922 23.9954C-0.0945812 22.5178 -0.452422 20.2016 0.6344 18.3191Z"
        fill={`url(#${g1})`}
      />
      <path
        d="M32.3651 0C34.9756 0.000564347 37.0926 2.11693 37.0926 4.72752V26.7138C37.0926 27.4836 37.926 27.9647 38.5927 27.5798L51.4483 20.1566L50.4788 18.4784C49.7749 17.2591 50.6015 15.7268 52.007 15.6461L66.8959 14.7943C68.423 14.707 69.4171 16.3738 68.6133 17.6751L60.661 30.5419C59.9089 31.7588 58.1308 31.7322 57.4154 30.4934L56.1758 28.3467L50.7688 31.4665C47.8438 33.1542 44.2407 33.1537 41.3161 31.4653L32.3653 26.2978C29.4398 24.6088 27.6376 21.4874 27.6376 18.1094V4.72752C27.6376 2.11658 29.7542 0 32.3651 0Z"
        fill={`url(#${g2})`}
      />
      <defs>
        {[g0, g1, g2].map((id) => (
          <linearGradient
            key={id}
            id={id}
            x1="47.8477"
            y1="52"
            x2="83.3477"
            y2="5"
            gradientUnits="userSpaceOnUse"
          >
            <stop stopColor="#2EB52E" />
            <stop offset="1" stopColor="#2EB52E" stopOpacity="0.51" />
          </linearGradient>
        ))}
      </defs>
    </svg>
  );
}

export function BrandLogo({ name }: { name: string }) {
  return (
    <a
      className="demo-brand"
      href="https://eigenpal.com"
      target="_blank"
      rel="noopener noreferrer"
      title={`${name} by EigenPal`}
    >
      <PptxIcon />
      <span className="demo-brand__name">{name}</span>
    </a>
  );
}
