const express = require('express');
const router = express.Router();
const docController = require('../controllers/docController');

router.post('/generate-doc', docController.generateDoc);
router.get('/download/:filename', docController.downloadDoc);
router.get('/', docController.homePage);

module.exports = router;