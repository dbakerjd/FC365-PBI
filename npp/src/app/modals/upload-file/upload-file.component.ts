import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from 'src/app/shared/forms/upload-file.config';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormGroup } from '@angular/forms';

@Component({
  selector: 'app-upload-file',
  templateUrl: './upload-file.component.html',
  styleUrls: ['./upload-file.component.scss']
})
export class UploadFileComponent implements OnInit {
  formConfig: UploadFileConfig = new UploadFileConfig();
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any; 

  constructor(@Inject(MAT_DIALOG_DATA) public data: any) { 
    
  }

  ngOnInit(): void {
    this.formConfig = new UploadFileConfig();
    this.fields = this.formConfig.fields(this.data.folderList);
    this.form = new FormGroup({});
  }

}
