function drawLabelField(label, name, x, y, width){

  const boxWidth = width + 60;
  const boxHeight = 20;

  page.drawRectangle({
    x: x,
    y: y - 6,
    width: boxWidth,
    height: boxHeight,
    borderWidth: 0.5,
    borderColor: rgb(0,0,0)
  });

  page.drawText(label,{
    x: x + 4,
    y: y,
    size: 9,
    font
  });

  const f = form.createTextField(name);
  f.setText(data[name] || "");
  f.setTextColor(rgb(0,0,0));
  f.setBorderWidth(0);

  f.addToPage(page,{
    x: x + 60,
    y: y - 5,
    width: width,
    height: 12
  });

}









































































