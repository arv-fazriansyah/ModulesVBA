async function handleRequest(request, env) {
  const url = new URL(request.url);
  const path = url.pathname;

  // Membuat pemetaan jalur ke URL target dari variabel lingkungan
  const pathToProxyUrl = {
    [`/${env.TOKEN}`]: env.TOKEN_URL,
    [`/${env.DATA}`]: env.DATA_URL,
    [`/${env.FORMULA}`]: env.FORMULA_URL,
  };

  // Memeriksa apakah jalur ada dalam pemetaan
  if (path in pathToProxyUrl) {
    const proxyUrl = pathToProxyUrl[path];
    return await fetch(proxyUrl);
  }

  // Mengembalikan respons 404 untuk jalur lainnya
  return new Response('Not Found', { status: 404 });
}

export default {
  // Fungsi untuk menangani permintaan fetch
  async fetch(request, env, ctx) {
    return await handleRequest(request, env);
  },
};
