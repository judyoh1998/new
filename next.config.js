/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    serverComponentsExternalPackages: ['pptxgenjs'],
  },
};

module.exports = nextConfig;
