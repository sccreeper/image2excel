const imageElement = document.querySelector("img");
const videoElement = document.getElementById("video");
const colourInput = document.getElementById("filter_colour");
const filterElement = document.getElementById("image_filter");

var size;
var boundingRect;

function changeFilter() {

  // filterElement.style.display = "block"
  
  if(currentMedia === "image") {
    size = [imageElement.width, imageElement.height];

    boundingRect = imageElement.getBoundingClientRect();

  } else {
    size = [videoElement.width, videoElement.height];

    boundingRect = videoElement.getBoundingClientRect();
  }
  
  console.log(boundingRect)
  console.log(size)

  filterElement.style.left = boundingRect.x + "px";
  filterElement.style.top = boundingRect.y + "px";
  
  filterElement.style.width = size[0] + "px"
  filterElement.style.height = size[1] + "px"
  
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