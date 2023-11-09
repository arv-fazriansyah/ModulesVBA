addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
});

async function handleRequest(request) {
  const url = new URL(request.url);
  const path = url.pathname;

  // Define a mapping of paths to JSON API URLs
  const pathToApiUrl = {
    '/token': 'YOUR_JSON_API_URL_TOKEN',
    '/data': 'YOUR_JSON_API_URL_DATA',
    '/info': 'YOUR_JSON_API_URL_INFO',
    '/another': 'YOUR_JSON_API_URL_ANOTHER',
    '/example': 'YOUR_JSON_API_URL_EXAMPLE',
  };

  // Check if the path is in the mapping
  if (path in pathToApiUrl) {
    const apiUrl = pathToApiUrl[path];
    return await handleJsonApiRequest(apiUrl);
  }

  // Return a 404 response for other paths
  return new Response('Not Found', { status: 404 });
}

async function handleJsonApiRequest(apiUrl) {
  const response = await fetch(apiUrl);
  const jsonData = await response.json();

  const htmlTable = generateHtmlTable(jsonData);

  return new Response(htmlTable, {
    headers: {
      'Content-Type': 'text/html',
    },
  });
}

function generateHtmlTable(jsonData) {
  let htmlTable = '<table border="1"><tr>';

  for (const key in jsonData[0]) {
    htmlTable += `<th>${key}</th>`;
  }

  htmlTable += '</tr>';

  for (const row of jsonData) {
    htmlTable += '<tr>';

    for (const key in row) {
      htmlTable += `<td>${row[key]}</td>`;
    }

    htmlTable += '</tr>';
  }

  htmlTable += '</table>';

  return htmlTable;
}
