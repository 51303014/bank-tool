const express = require("express");
const router = express.Router();
const exportFile = require("./export");
const loginUser = require("./login");
// const checkPassword = require("./check-password");
// const { tokenVerification } = require("../../middleware");

// ROUTES * /api/auth/
router.post("/login", loginUser);
router.post("/export", exportFile);

module.exports = router;
