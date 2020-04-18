function update_image(){
    var img = document.querySelector('img'); // Returns the first img element
    var file = document.querySelector('input[type=file]').files[0]; // Returns the first file element found
    img.src =  window.URL.createObjectURL(file);

}
