async function handleRequest(request, env) {
  const url = new URL(request.url);
  const path = url.pathname;

  const pathToProxyUrl = {
    [`/${env.DATA}`]: env.DATA_URL,
    [`/${env.FORMULA}`]: env.FORMULA_URL,
    [`/${env.DEV}`]: env.DEV_URL,
    [`/${env.SEND_DATA}`]: env.SEND_DATA_URL,
  };

  const proxyUrl = pathToProxyUrl[path];
  if (proxyUrl) {
    return await fetch(proxyUrl, request);
  }

  return new Response('Not Found', { status: 404 });
}

export default {
  async fetch(request, env, ctx) {
    return await handleRequest(request, env);
  },
};
