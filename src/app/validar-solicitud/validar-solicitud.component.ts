import { Component, OnInit } from '@angular/core';
import { ItemAddResult, EmailProperties } from '@pnp/sp';
import { Validators, FormBuilder, FormGroup } from '@angular/forms';
import { SPServicio } from '../Servicios/sp-servicio';
import { Ng4LoadingSpinnerService } from 'ng4-loading-spinner';
import { NgxSpinnerService } from 'ngx-spinner';
import { ToastrManager } from 'ng6-toastr-notifications';
import { Usuario } from '../Entidades/usuario';
import { Solicitud } from '../Entidades/solicitud';
import { debug } from 'util';
import { ActivatedRoute } from '@angular/router';
import SimpleCrypto from "simple-crypto-js";
import * as CryptoJS from 'crypto-js';

@Component({
  selector: 'app-validar-solicitud',
  templateUrl: './validar-solicitud.component.html',
  styleUrls: ['./validar-solicitud.component.css']
})
export class ValidarSolicitudComponent implements OnInit {

  SolicitudPermisoForm: FormGroup;
  usuarioActual: Usuario;
  otroTipoPermiso: boolean;

  hora: string[] = [ "00:00", "00:30", "01:00", "01:30", "02:00", "02:30", "03:00", "03:30", "04:00", "04:30", "05:00", "05:30", 
  "06:00", "06:30", "07:00", "07:30", "08:00", "08:30", "09:00", "09:30", "10:00","10:30", "11:00", "11:30", "12:00", "12:30", 
  "13:00", "13:30", "14:00", "14:30", "15:00", "15:30", "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", 
  "19:00", "19:30", "20:00", "20:30", "21:00", "21:30", "22:00", "22:30", "23:00", "23:30"];
  TipoPermiso: any;
  minDateFechaInicio: Date;
  minDateFechaFin: Date;
  ObjSolicitud: Solicitud[];
  AprobarJefe: boolean;
  idSolicitud: string;
  IdServicio: any;
  RecepcionGH: boolean;
  EmailSolicitante: string;
  UsuarioHG: boolean;
  ComentarioGH: boolean;
  MensajeAccion: string;
  objUsuariosGH: any[]=[];
  objUsuariosGHcorreos: any[]=[];

  constructor( 
    private formBuilder: FormBuilder, 
    public toastr: ToastrManager,     
    private spinnerService: Ng4LoadingSpinnerService,
    private servicio: SPServicio,
    private route: ActivatedRoute) { 
    this.otroTipoPermiso = false;
    this.TipoPermiso = "";
    this.minDateFechaInicio = new Date();
    this.AprobarJefe = false;
    this.RecepcionGH = false;
    this.ComentarioGH = false;
  }

  ngOnInit() {
    this.spinnerService.show();     
    
    this.SolicitudPermisoForm = this.formBuilder.group({
      Apellidos: [''],
      Nombres: [''],
      Cedula: [''],
      CargoRol: [''],
      UnidadNegocio: [''],
      TelefonoExt: [''],
      TipoPermiso: ["", Validators.required],
      OtroMotivo: [""],
      MotivoPermiso: ["", Validators.required],
      HoraInicio: ["", Validators.required],
      FechaInicio: ["", Validators.required],
      HoraFin: ["", Validators.required],
      FechaFin: ["", Validators.required],
      ObservacionGH: [""]
    });      
    
    this.ObtenerUsuarioActual();
  }

  ObtenerUsuarioActual() {
    this.servicio.ObtenerUsuarioActual().subscribe(
      (respuesta) => {
        this.usuarioActual = new Usuario(respuesta.Id);        
        this.ValidarUsuarioGH()
      }, err => {
        console.log('Error obteniendo usuario: ' + err);
      }
    )
  }

  ValidarUsuarioGH(){
    this.servicio.ObtenerUsuarioGH().then(
      (res)=>{
        let objUsuariosGH = [];
        

        this.objUsuariosGH = res[0].GestionHumanaId;
        this.objUsuariosGHcorreos = res[0].GestionHumana;
        let usuarioGH = this.objUsuariosGH.filter(x=> x === this.usuarioActual.idUsuario);
        if (usuarioGH.length > 0) {
            this.UsuarioHG = true;
        }
        else{
          this.UsuarioHG = false;
        }
        this.consultarSolicitudPermiso();
      }
    ).catch(
      (error)=>{
        console.log(error);
        this.mostrarError("Error al cargar las solicitudes");
        this.spinnerService.hide();
      }
    ) 
  }

  getParams(url){
    let params = {};
  	let parser = document.createElement('a');
  	parser.href = url;
  	let query = parser.search.substring(1);
  	let vars = query.split('&');
  	for (let i = 0; i < vars.length; i++) {
  		let pair = vars[i].split('=');
  		params[pair[0]] = decodeURIComponent(pair[1]);
  	}
  	return params;
  }

  consultarSolicitudPermiso(): any {    

    let Parametro = this.getParams(window.location.href);    
    this.idSolicitud = Parametro["id"].substring(12, Parametro["id"].length);

    if (this.idSolicitud === null || this.idSolicitud === "") {
      this.mostrarError("Error al cargar la solicitud");
      setTimeout(
        ()=>{
          this.spinnerService.hide();
        },2000);
    }
    this.servicio.consultarSolicitudPermiso(this.idSolicitud).then(
      (respuesta)=>{        
        this.ObjSolicitud = Solicitud.fromJsonList(respuesta);
        this.SolicitudPermisoForm.controls["TipoPermiso"].setValue(this.ObjSolicitud[0].tipoPermiso);
        this.SolicitudPermisoForm.controls["OtroMotivo"].setValue(this.ObjSolicitud[0].otroTipoPermiso);      
        this.SolicitudPermisoForm.controls["MotivoPermiso"].setValue(this.ObjSolicitud[0].motivoPermiso);
        this.SolicitudPermisoForm.controls["HoraInicio"].setValue(this.ObjSolicitud[0].horaInicioPermiso);
        this.SolicitudPermisoForm.controls["FechaInicio"].setValue(this.ObjSolicitud[0].fechaInicioPermiso);
        this.SolicitudPermisoForm.controls["HoraFin"].setValue(this.ObjSolicitud[0].horaFinPermiso);
        this.SolicitudPermisoForm.controls["FechaFin"].setValue(this.ObjSolicitud[0].fechaFinPermiso);
        this.SolicitudPermisoForm.controls["ObservacionGH"].setValue(this.ObjSolicitud[0].ObservacionGH);
        this.EmailSolicitante = this.ObjSolicitud[0].EmailSolicitante;
        let actualUser = this.ObjSolicitud[0].responsableActual.filter(x=> x === this.usuarioActual.idUsuario);
        console.log(actualUser);
        if (this.ObjSolicitud[0].tipoPermiso === "Otro") {
          this.otroTipoPermiso = true;
        }
        if (this.ObjSolicitud[0].estado === "En revision GH" && this.UsuarioHG === true) {
          this.RecepcionGH = true;
          this.ComentarioGH = true;                      
        }

        if (this.ObjSolicitud[0].estado === "Recibido por GH") {          
            this.ComentarioGH = true;
            this.SolicitudPermisoForm.controls["ObservacionGH"].disable();
        }
        
        if (actualUser.length > 0) {

          if (this.ObjSolicitud[0].estado === "En revision jefe" && (this.usuarioActual.idUsuario === actualUser[0])) {
            this.AprobarJefe = true;
          }
        }
        
        this.ObtenerSolicitante(this.ObjSolicitud[0].idUsuario);
      }
    ).catch(
      (error)=>{
        console.log(error);
        this.mostrarError("Error al cargar la solicitud");
      }
    );    
  }

  ObtenerSolicitante(idSolicitante: any): any {
    this.servicio.obtenerUsuarioListaEmpleados(idSolicitante).then(
      (res)=>{
        
        if (res.length > 0) { 
                              
          let SegundoNombre = res[0].SegundoNombre === null? "" : " "+res[0].SegundoNombre;
          let SegundoApellido = res[0].SegundoApellido === null? "" : " "+res[0].SegundoApellido;
          let nombre = res[0].PrimerNombre + SegundoNombre;
          let apellidos = res[0].PrimerApellido + SegundoApellido;
          let cedula = res[0].NumeroDocumento === null? "" :res[0].NumeroDocumento;
          let cargo = res[0].Cargo === null? "" : res[0].Cargo;
          let Area = res[0].Area === null? "" : res[0].Area;
          let celular = res[0].Celular === null? "" : res[0].Celular;

          this.SolicitudPermisoForm.controls["Apellidos"].setValue(apellidos);
          this.SolicitudPermisoForm.controls["Nombres"].setValue(nombre);
          this.SolicitudPermisoForm.controls["Cedula"].setValue(cedula);
          this.SolicitudPermisoForm.controls["CargoRol"].setValue(cargo);
          this.SolicitudPermisoForm.controls["UnidadNegocio"].setValue(Area);
          this.SolicitudPermisoForm.controls["TelefonoExt"].setValue(celular);
          
          this.servicio.ObtenerServicio(this.idSolicitud).then(
            (resServicio)=>{
              this.IdServicio = resServicio[0].Id;
              this.spinnerService.hide();
            }            
          ).catch(
            (error)=>{
              console.log(error);
              this.mostrarError("Error al obtener la solicitud");
              // window.location.reload();
            }
          );          
        }        
      }
    ).catch(
      (error)=>{
          console.log(error);
      }
    );
  }
  
  MostrarExitoso(mensaje: string) {
    this.toastr.successToastr(mensaje, 'Confirmación!');
  }

  mostrarError(mensaje: string) {
    this.toastr.errorToastr(mensaje, 'Oops!');
  }

  mostrarAdvertencia(mensaje: string) {
    this.toastr.warningToastr(mensaje, 'Validación');
  }

  mostrarInformacion(mensaje: string) {
    this.toastr.infoToastr(mensaje, 'Información importante');
  }
   
  AprobarSolicitud(){
    this.spinnerService.show(); 
    let fecha = new Date();
    let hora = fecha.toLocaleTimeString(navigator.language, {hour: '2-digit', minute:'2-digit'});

    let ObjRespuestaJefe = {
      HoraAprobacionLider: hora,
      FechaAprobacionLider: fecha,
      AprobacionLider: true,
      Estado: "En revision GH",
      ResponsableActualId:  {
        results: this.objUsuariosGH
      }
      
    }

    
    this.servicio.GuardarRespuestaJefe(ObjRespuestaJefe, this.idSolicitud).then(
      (itemResult)=>{
        let objServicio = {          
          ResponsableActualId: this.objUsuariosGH[0],
          Estado: "En revision GH"
        }
        
        this.MensajeAccion = "La solicitud se ha aprobado con éxito";
        this.enviarNotificacion(objServicio, "Aprobada");
        // this.guardarServicio(objServicio);   
      }
    ).catch(
      (error)=>{
        console.log(error);
        this.mostrarError("Error al aprobar la solicitud");        
      }
    );
  }
  
  RechazarSolicitud(){
    let fecha = new Date();
    let hora = fecha.toLocaleTimeString(navigator.language, {hour: '2-digit', minute:'2-digit'});

    let ObjRespuestaJefe = {
      HoraAprobacionLider: hora,
      FechaAprobacionLider: fecha,
      AprobacionLider: false,
      Estado: "Rechazado",
      ResponsableActualId: {
        results: [0]
      },
    }

    this.MensajeAccion = "La solicitud ha sido rechazada con éxito";
    this.servicio.GuardarRespuestaJefe(ObjRespuestaJefe, this.idSolicitud).then(
      (itemResult)=>{
        let objServicio = {          
          ResponsableActualId: -1,
          Estado: "Rechazado"
        }
        this.enviarNotificacion(objServicio, "Rechazada");
        // this.guardarServicio(objServicio);   
      }
    ).catch(
      (error)=>{
        console.log(error);
        this.mostrarError("Error al aprobar la solicitud");        
      }
    );
  }

  RecibirSolicitud(){
    console.log("jam");
    let fecha = new Date();
    let hora = fecha.toLocaleTimeString(navigator.language, {hour: '2-digit', minute:'2-digit'});
    let observacionGH = this.SolicitudPermisoForm.controls["ObservacionGH"].value;
    let ObjRespuestaJefe = {
      HoraRecepcionGH: hora,
      FechaRecepcionGH: fecha,
      Estado: "Recibido por GH",
      ResponsableActualId: {
        results: [0]
      },
      ObservacionGH: observacionGH
    }
    this.MensajeAccion = "La solicitud ha sido recibida con éxito";
    this.servicio.GuardarRecepcionGH(ObjRespuestaJefe, this.idSolicitud).then(
      (itemResult)=>{
        let objServicio = {          
          ResponsableActualId: -1,
          Estado: "Recibido por GH"
        }
        
        this.guardarServicio(objServicio);
      }
    ).catch(
      (error)=>{
        console.log(error);
        this.mostrarError("Error al aprobar la solicitud");        
      }
    );
  }
    
  enviarNotificacion(objServicio, accion): any {
    let correos = "";
          let TextoCorreo = '<p>Cordial saludo</p>'+
                            '<br>'+
                            '<p>Se le informa que la solicitud de permiso ha sido '+accion+'</p>';

          this.objUsuariosGHcorreos.forEach(element => {
            
            correos += element.EMail +";";
               
            console.log(correos);
          });

          const emailProps: EmailProperties = {
            To: correos.split(';'),
            CC: [this.EmailSolicitante],
            Subject: "Notificación de solicitud de permiso",
            Body: TextoCorreo,
          };
          this.servicio.EnviarNotificacion(emailProps).then(
            (res)=>{
                  this.guardarServicio(objServicio);                  
            }
          ).catch(
            (error)=>{
              this.mostrarError("Error al enviar la notificacion");
              setTimeout(
                ()=>{
                  window.location.href = 'https://aribasas.sharepoint.com/sites/Intranet';
                  this.spinnerService.hide();
                },2000);
              
            }
          ) 
  }

  guardarServicio(objServicio) {
    this.servicio.ModificarServicio(objServicio, this.IdServicio).then(
        (resultado)=>{
          this.MostrarExitoso(this.MensajeAccion);
          sessionStorage.removeItem("TipoConsulta");
          setTimeout(
            ()=>{
              window.location.href = 'https://aribasas.sharepoint.com/sites/Intranet';
              this.spinnerService.hide();
            },2000);
        }
    ).catch(
      (error)=>{
        console.log(error);
        this.mostrarError("Error al aprobar la solicitud");
        setTimeout(
          ()=>{
            window.location.href = 'https://aribasas.sharepoint.com/sites/Intranet';
            this.spinnerService.hide();
          },2000);
      }
    );
  }


}
