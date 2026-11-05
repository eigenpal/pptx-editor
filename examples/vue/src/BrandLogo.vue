<script setup lang="ts">
// The "PPTX Editor by EigenPal" lockup: the Vue twin of examples/vite/src/BrandLogo.tsx, with
// the same presentation slide icon and the same `demo-brand` styles in examples/shared/shell.css.
defineProps<{ name: string }>();
</script>

<template>
  <a
    class="demo-brand"
    href="https://eigenpal.com"
    target="_blank"
    rel="noopener noreferrer"
    :title="`${name} by EigenPal`"
  >
    <svg
      viewBox="0 0 100 84"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden="true"
      style="width: 30px; height: auto; flex-shrink: 0"
    >
      <path
        d="M2 2H98V60H2V2Z"
        fill="#FAF9F8"
        stroke="#201F1E"
        stroke-width="4"
        stroke-linejoin="round"
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
      <rect x="28" y="66" width="44" height="16" fill="#C43E1C" stroke="#201F1E" stroke-width="3" />
      <text
        x="50"
        y="78"
        text-anchor="middle"
        fill="#FAF9F8"
        font-size="11"
        font-weight="900"
        font-family="system-ui, sans-serif"
      >
        .PPTX
      </text>
    </svg>
    <span class="demo-brand__name">{{ name }}</span>
  </a>
</template>
