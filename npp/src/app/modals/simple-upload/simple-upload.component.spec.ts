import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SimpleUploadComponent } from './simple-upload.component';

describe('SimpleUploadComponent', () => {
  let component: SimpleUploadComponent;
  let fixture: ComponentFixture<SimpleUploadComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SimpleUploadComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(SimpleUploadComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
