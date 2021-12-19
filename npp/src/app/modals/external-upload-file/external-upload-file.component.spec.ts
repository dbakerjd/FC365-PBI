import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExternalUploadFileComponent } from './external-upload-file.component';

describe('ExternalUploadFileComponent', () => {
  let component: ExternalUploadFileComponent;
  let fixture: ComponentFixture<ExternalUploadFileComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExternalUploadFileComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExternalUploadFileComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
