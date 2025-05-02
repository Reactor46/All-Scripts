import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { HttpModule, JsonpModule } from '@angular/http';
import { ModalModule } from "ng2-bootstrap";


import { AppComponent } from './app.component';
import { routing, appRoutingProviders } from './app.routing';
import { HomeComponent } from './home/home.component';

import { DataService } from "./services/data.service";

import { FileSizePipe } from "./pipes/file-size.pipe"
import { EventIDURL } from "./pipes/eventidurl.pipe"
import { NumFormatPipe } from "./pipes/numformat.pipe"

@NgModule({
    declarations: [
        FileSizePipe,
        EventIDURL,
        NumFormatPipe,
        AppComponent,
        HomeComponent,
    ],
    imports: [
        ModalModule.forRoot(),
        BrowserModule,
        FormsModule,
        HttpModule,
        JsonpModule,
        routing
    ],
    providers: [appRoutingProviders, DataService],
    bootstrap: [AppComponent]
})

export class AppModule {
}
