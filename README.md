# wordGPT
WordGPT - это скрипт, который работает с искусственным интеллектом и позволяет создавать высококачественные тексты прямо в документе Word используя API OpenAI.

### Как установить?
---
1. Зайдите в Word и выберите **Макросы**

![image](https://user-images.githubusercontent.com/78819104/223409119-add85188-419a-467e-a1f3-bde8629f6609.png)

2. Создайте макрос с именем **AddToShortcut** и нажмите **Изменить**

![image](https://user-images.githubusercontent.com/78819104/223406919-a902d0cc-5e49-4855-9485-371530aa8325.png)


3. Вставьте код [ask.bas](ask.bas) 

4. Включите *Microsoft Scripting Runtime* в меню **Tools - References** 

![image](https://user-images.githubusercontent.com/78819104/223406090-a6c0459c-0cdd-489b-8c30-e0be3ae0bd56.png)

![image](https://user-images.githubusercontent.com/78819104/223406231-fe599510-32b7-456c-9837-83defc1e54ab.png)

5. Выполните макрос **AddToShortcut**

![image](https://user-images.githubusercontent.com/78819104/223406919-a902d0cc-5e49-4855-9485-371530aa8325.png)

6. Напишите желаемый запрос на странице в Word, нажмите правой кнопкой мыши по выделенному предложению и нажмите **Спросить ChatGPT**

![image](https://user-images.githubusercontent.com/78819104/223407377-fe290c0e-0421-4b29-a47f-82a352eca1d4.png)

## Результат ✨
---

![image](https://user-images.githubusercontent.com/78819104/223408081-f8896f5a-ba69-4980-9858-536d0853873c.png)

*Чтобы удалить "Спросить ChatGPT" в контекстном меню, зайдите в макросы и выполните **DeleteShortcut***
