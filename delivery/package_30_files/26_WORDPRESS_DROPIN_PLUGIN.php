<?php
/**
 * Plugin Name: Excel Arşiv Universal AI Search Drop-in
 * Description: Edge AST Pruning and /llms.txt headers.
 * Version: 3.0.0
 */
add_action('init', function() {
  if (strpos($_SERVER['REQUEST_URI'], '/llms.txt') !== false) {
    header('Content-Type: text/markdown; charset=utf-8');
  }
});
