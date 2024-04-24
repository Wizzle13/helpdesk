const express = require('express');
const app = express();
const exphbs = require('express-handlebars');
const hbs = exphbs.create();
const path = require('path');

const PORT = process.env.PORT || 5000;

// set Handlebars Middleware
app.engine('handlebars',  hbs.engine);
app.set('view engine', 'handlebars');

// Set Handlebars Rouets
app.get('/', function (req, res) {
    res.render('home');
})

// Set static folder
app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => console.log('Server Listening on port ' + PORT));