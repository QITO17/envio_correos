import csv
import datetime
import win32com.client as win32
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

#Correos Masivos Por Outlook o como se llame xd

# !   ****************************** Lea Atentamente Los Comentarios Para Manipular el Codigo ***************************************  !
#TODO AUTHORES JOSTIN ARLEY HURTADO -  JHON EDUARD OCAMPO

#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=>
#? =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> =>★ => ★=> ★=> ★=> ★=> ★=>★ => ★=> ★=> 


#! INICIO FUNCION ENVIAR CORREOS

def envia_correos():
    try:
        #archivo = open(valor_rescatado, "w")
        #archivo.close()
        global asunto_entry_get
        global cuerpo_entry_get
        asunto_entry_get = asunto_entry.get()
        cuerpo_entry_get = cuerpo_entry.get("1.0", "end")

        # with open(valor_rescatado, "a+", newline="", encoding="utf-8") as file:
        #     write = csv.writer(file, delimiter=",", quoting=csv.QUOTE_MINIMAL)
            # write.writerow([
            #                     'Correo',	'Asunto',	'Texto'
            #                 ])
            
        with open(valor_rescatado, encoding="utf-8") as File:
            reader = csv.reader(File, delimiter=",", quoting=csv.QUOTE_MINIMAL)
            maestra = list(reader)
        print(maestra)   

        i=0
        no_enviados = []
        for correo in maestra[1:]:
            corr = correo[0]

            if corr != "":
                outlook = win32.Dispatch("Outlook.Application")

                print('hola32')

                mail = outlook.CreateItem(0)

                print(mail)

                print(corr)

                mail.To = corr

                mail.Subject = asunto_entry_get

                mail.body = cuerpo_entry_get

                #adjunto = "C:/Users/CoorSistemas/Desktop/Correos/ejemplo.txt"  # Ruta al archivo que deseas adjuntar

                attachment = mail.Attachments.Add(valor_rescatado2)

                mail.Send()

                try:
                    mail.Send()
                except Exception as e:
                    no_enviados.append(corr)   
            else:
                        
                estado = "Correo No Enviado"

                print('hola32')

                print('hola el correo no esta')


            print(i)
            i=i+1
        # if len(no_enviados) > 0:
        #     messagebox.showinfo("Mensaje Informativo", "Los siguientes correos no se enviaron.")
        #     mensaje = "El contenido de la lista es:\n" + ", ".join(map(str, no_enviados))
        #     messagebox.showinfo("Correos no enviados", mensaje)

        limpia_campos()
        
    except Exception as error:
        messagebox.showinfo("Correos no enviados", error)

#! FIN FUNCION ENVIAR CORREOS


#! Limpiar Campos
def limpia_campos():
    asunto_entry.delete(0, "end")
    cuerpo_entry.delete("1.0", "end")
    messagebox.showinfo("Mensaje Informativo", "Correos enviados con exito.")


#! Funcion Rescatar Valor EXCEL E IMG
    
def rescata_url_csv():
    global valor_rescatado
    valor_rescatado = cargar_archivo()

def rescata_url_img():
    global valor_rescatado2
    valor_rescatado2 = cargar_img()


#! FUNCION CARGAR ARCHIVOS

def cargar_archivo():
    arch = filedialog.askopenfilename(title="Abrir")
    print('ex ', arch)
    messagebox.showinfo("Mensaje Informativo", "Archivo cargado.")

    return arch

def cargar_img():
    arch = filedialog.askopenfilename(title="Abrir")
    print('img ', arch)
    messagebox.showinfo("Mensaje Informativo", "Imagen cargada.")
    return arch
#! FIN CARGAR ARCHIVOS



  



color_fondo = "#f0f0f0"  # Color de fondo para la ventana
color_titulo = "#4caf50"  # Color del texto del título
color_labels = "#333333"  # Color del texto de los labels
color_botones = "#4caf50"  # Color de fondo de los botones
color_texto_botones = "white"  # Color del texto de los botones

#Configuración de la ventana Tamaño. Que no se pueda estirar ni encoger, Titulo Y un texto principal
#Con Mainloop mantengo la pertaña siempre activa es como un while infinito
mi_ventana = Tk()
RUTA_IMG = ''
RUTA_EX = ''
mi_ventana.geometry("650x550")
mi_ventana.title("Envio correo masivos")
mi_ventana.resizable(False, False)
titulo_principal = Label(text="Envio de correos Cootransmede", font=("Cambria", 18), width="550", height="2")
titulo_principal.pack()
#mi_ventana.iconbitmap("cmt2.ico")






# * Aqui estan los label del programa podra apreciar que estan 2 veces uno usa el metodo label y el otro el metodo place
# * El de label es para poer el texto y el que usa el metodo place es para darle posición en (X) y (Y) en la ventana

label_asunto = Label(text="Asunto", font=("Cambria", 13), height="2")
label_asunto.place(x=30, y=100)

label_cuerpo = Label(text="Cuerpo correo", font=("Cambria", 13), height="2")
label_cuerpo.place(x=30, y=140)

asunto = StringVar()
cuerpo = StringVar()

asunto_entry = Entry(textvariable=asunto, width="71")
cuerpo_entry = Text(width=54)

boton_submit = Button(text="Cargar CSV", padx=37, pady=1, command=rescata_url_csv)
boton_submit.place(x=100, y=600)
boton_submit.pack(side=BOTTOM, pady = 10, padx = 25)

boton_envia_correo = Button(text="Enviar Correos", padx=25, pady=1, command=envia_correos)#
boton_envia_correo.place(x=100, y=600)
boton_envia_correo.pack(side=BOTTOM, pady = 10, padx = 25)

boton_envia_img = Button(text="Cargar Imagen", padx=25, pady=1, command=rescata_url_img)
boton_envia_img.place(x=100, y=600)
boton_envia_img.pack(side=BOTTOM, pady = 10, padx = 25)

scrll = Scrollbar(mi_ventana, command=cuerpo_entry.yview)
scrll.place(x=580, y=155, height=100)


asunto_entry.place(x=150, y=115, height=20)
cuerpo_entry.place(x=150, y=155, height=100)


mi_ventana.config(bg=color_fondo)
titulo_principal.config(bg=color_titulo, fg="#fff") 
boton_submit.config(bg=color_botones, fg="white", font=("Cambria", 13))
boton_envia_correo.config(bg=color_botones, fg="white", font=("Cambria", 13))
boton_envia_img.config(bg=color_botones, fg="white", font=("Cambria", 13))

mi_ventana.mainloop()
