<?php

namespace zkyg\SpreadsheetFormulaLaravel;
use Illuminate\Support\ServiceProvider;

class SpreadsheetFormulaServiceProvider extends ServiceProvider
{

    function boot(){
    }

    function register()
    {
        $this->app->singleton(SpreadsheetFormulaParser::class, function (){
            return new SpreadsheetFormulaParser();
        });
    }
}
