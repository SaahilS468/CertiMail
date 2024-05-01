# CertiMail
This is a program for creating certificates from a word document and mail it to the respective email id's.<br/>
### The requirements before running the code:
- A Word document containing the template of the certificate
    - Make sure the document is prepared in such a way that the template covers the whole document space(such as landscape).
    - Add a placeholder to populate names in the black field on the template as below <br/> `{{Variable used in the code}}`.
- A CSV file containing the **Names** and **Emails** of the users.
- Mention the path to where the certificates needs to be created.
- Create your app password from your google account using which the email will be sent and add both the password and email id in `.env` file.

### Download the required libraries: </br>
`pip install -r requirements.txt`