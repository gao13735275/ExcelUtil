function uploadPic() {
    var form = document.getElementById('upload'),
        formData = new FormData(form);
    $.ajax({
        url:"/importExcel",
        type:"post",
        data:formData,
        processData:false,
        contentType:false,
        success:function(res){
            alert("文件已上传！");
            console.log(res);
        }

    })

}