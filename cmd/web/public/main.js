// Form elements

const file_input = document.getElementById("file_input")
const file_image = document.getElementById("file_image")
const scale_input = document.getElementById("image_scale")
const scale_label = document.getElementById("image_scale_label")

const convert_button = document.getElementById("convert_button")
const status_label = document.getElementById("status")

// Other constants

const max_width = 500

scale_input.addEventListener("input", () => {
    scale_label.textContent = `Scale: ${scale_input.value}`
})

file_input.addEventListener("change", () => {

    file_image.setAttribute("src", URL.createObjectURL(file_input.files[0]))

})

file_image.addEventListener("load", () => {
    if (file_image.clientWidth > max_width) {
        file_image.setAttribute("width", max_width)
    }
})

convert_button.addEventListener("click", () => {

    status_label.textContent = "Converting..."

    var data = new FormData()
    data.append("file", file_input.files[0]) 
    data.append("scale", parseFloat(scale_input.value))

    console.log(parseFloat(scale_input.value))

    fetch("/convert", {
        method : "POST",
        body: data
    }).then( (response) => response.arrayBuffer()).then((data) => {

        status_label.textContent = "Finished."

        let blob = new Blob([data], {type : "application/vnd.ms-excel"})

        let link = document.createElement("a")
        link.href = URL.createObjectURL(blob)

        link.download = `${file_input.files[0].name}.xlsx`

        link.click();

    })

})