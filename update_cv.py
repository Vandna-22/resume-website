import docx

doc = docx.Document('Vandna_Gupta_CV_FINAL.docx')
runs = doc.paragraphs[2].runs
runs[0].text = 'Khurja, Uttar Pradesh, India   |   vandnaa2230@gmail.com   |   🌐 www.linkedin.com/in/vandna-gupta-6a579a351   |   💻 github.com/Vandna-22'
doc.save('Vandna_Gupta_CV_FINAL.docx')
print("CV updated successfully")
