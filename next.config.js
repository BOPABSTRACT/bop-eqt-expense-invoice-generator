/** @type {import('next').NextConfig} */
const nextConfig = {
  typescript: { ignoreBuildErrors: true },
  eslint: { ignoreDuringBuilds: true },
  experimental: {
    serverComponentsExternalPackages: ['jspdf', 'jspdf-autotable', 'pdf-lib']
  }
}

module.exports = nextConfig
