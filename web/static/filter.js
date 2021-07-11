const imageElement = document.querySelector("img");
const colourInput = document.getElementById("filter_colour")
const filterElement = document.getElementById("image_filter")

const boundingRect = imageElement.getBoundingClientRect();

filterElement.style.left = boundingRect.x;
filterElement.style.top = boundingRect.y;

function changeFilter() {

  
  var size = [imageElement.width, imageElement.height];

  filterElement.style.width = size[0]
  filterElement.style.height = size[1]
  
  var colour = hexToRgb(colourInput.value);

  filterElement.style.backgroundColor = `rgba(${colour.r}, ${colour.g}, ${colour.b}, 0.5)`

}

//https://stackoverflow.com/a/5624139
function hexToRgb(hex) {
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}