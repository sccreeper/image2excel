const url = '/convert'
const form = document.getElementById('submit')

form.addEventListener('click', e => {
  e.preventDefault()
    
  var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
      if (this.readyState == 4 && this.status == 200) {
        loop(this.responseText);
      }
    };

    var form = document.getElementById('form');
    var form_data = new FormData(form);

    xhttp.open("POST", "/convert", true);
    xhttp.send(form_data);
  
});

var i = 0;

function loop(pid) {  
  setTimeout(function () {           
    i++;                     
      // rest of code here
      var xhttp = new XMLHttpRequest();
      xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
        //document.getElementById("demo").innerHTML = this.responseText;
  
         var data = JSON.parse(this.responseText);
         document.getElementById('progress_bar').style.width = data.progress + 'px';
         document.getElementById('progress_bar').setAttribute('title', data.progress + '% complete');

         document.getElementById('progress_status').innerHTML = data.status;
  
         if (data.finished === 'True') {
           finished = true;
           window.open('/convert/' + pid, '_blank');
         } else {
           loop(pid);
         }  
        };
      };
      xhttp.open("GET", "/convert/progress/" + pid, true);
      xhttp.send();                    
 }, 250)
  
}

function updateForm() {
  document.getElementById('file_b').innerHTML = document.getElementById('img_file')[0].value.split(/(\\|\/)/g).pop();
}

function update_image(){
  var img = document.getElementById('image_file'); // Returns the first img element
  var file = document.querySelector('input[type=file]').files[0]; // Returns the first file element found
  img.src =  window.URL.createObjectURL(file);

}

function filterDetect(params) {
  var e = document.getElementById('filters')
  if(e.options[e.selectedIndex].value == 'FILTER') {
    document.getElementById('filter_colour').style.visibility = 'visible';
  } else {
    document.getElementById('filter_colour').style.visibility = 'hidden';
  }
}