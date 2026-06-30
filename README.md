Este repositorio es auxiliar para el envio de correos electronicos de manera automatizada, se estructura como modulo para poder ser utilizado en otros repositorios como submodulo

Para que el proceso se ejecute correctamente y de manera segura se recomienda definir las siguientes variables localmente en un archivo .env o en secrets para ejecucion en Github actions:

cuenta: correo de donde saldra el correo,

password: contraseña de la cuenta anterior (es necesaria pues la creacion del correo es mediante la web)

