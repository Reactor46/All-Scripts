import { Component, OnInit } from '@angular/core';
import "font-awesome/css/font-awesome.css"
import '../css/main.css';

@Component({
    selector: 'app',
    template: require('./app.component.html')
})

export class AppComponent implements OnInit {
    ngOnInit() {
        console.log('Initializing...');
    }
}
