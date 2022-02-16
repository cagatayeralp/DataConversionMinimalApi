A Simple Service coded with .net Core Minimal WebApi. This service has 2 endpoints

convert_excel : This endpoint takes a list of jsons and converts them into excel file
convert_pdf : This endpoint takes a list of jsons and convers them into a pdf file
Sample Input :

{
  "json_input": [
    {
      "productId": 1,
      "name": "Samosa",
      "description": "Praesent scelerisque, mi sed ultrices condimentum, lacus ipsum viverra massa, in lobortis sapien eros in arcu. Quisque vel lacus ac magna vehicula sagittis ut non lacus.<br/>Sed volutpat tellus lorem, lacinia tincidunt tellus varius nec. Vestibulum arcu turpis, facilisis sed ligula ac, maximus malesuada neque. Phasellus commodo cursus pretium.",
      "price": 15,
      "categoryName": "Appetizer",
      "imageUrl": "https://cagataydotnetmastery.blob.core.windows.net/mango/14.jpg"
    },
    {
      "productId": 2,
      "name": "Paneer Tikka",
      "description": "Praesent scelerisque, mi sed ultrices condimentum, lacus ipsum viverra massa, in lobortis sapien eros in arcu. Quisque vel lacus ac magna vehicula sagittis ut non lacus.<br/>Sed volutpat tellus lorem, lacinia tincidunt tellus varius nec. Vestibulum arcu turpis, facilisis sed ligula ac, maximus malesuada neque. Phasellus commodo cursus pretium.",
      "price": 13.99,
      "categoryName": "Appetizer",
      "imageUrl": "https://cagataydotnetmastery.blob.core.windows.net/mango/12.jpg"
    },
    {
      "productId": 3,
      "name": "Sweet Pie",
      "description": "Praesent scelerisque, mi sed ultrices condimentum, lacus ipsum viverra massa, in lobortis sapien eros in arcu. Quisque vel lacus ac magna vehicula sagittis ut non lacus.<br/>Sed volutpat tellus lorem, lacinia tincidunt tellus varius nec. Vestibulum arcu turpis, facilisis sed ligula ac, maximus malesuada neque. Phasellus commodo cursus pretium.",
      "price": 10.99,
      "categoryName": "Dessert",
      "imageUrl": "https://cagataydotnetmastery.blob.core.windows.net/mango/11.jpg"
    },
    {
      "productId": 4,
      "name": "Pav Bhaji",
      "description": "Praesent scelerisque, mi sed ultrices condimentum, lacus ipsum viverra massa, in lobortis sapien eros in arcu. Quisque vel lacus ac magna vehicula sagittis ut non lacus.<br/>Sed volutpat tellus lorem, lacinia tincidunt tellus varius nec. Vestibulum arcu turpis, facilisis sed ligula ac, maximus malesuada neque. Phasellus commodo cursus pretium.",
      "price": 15,
      "categoryName": "Entree",
      "imageUrl": "https://cagataydotnetmastery.blob.core.windows.net/mango/13.jpg"
    }
  ]
}