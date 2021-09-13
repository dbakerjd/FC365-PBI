import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExpiredLicenseComponent } from './expired-license.component';

describe('ExpiredLicenseComponent', () => {
  let component: ExpiredLicenseComponent;
  let fixture: ComponentFixture<ExpiredLicenseComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExpiredLicenseComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExpiredLicenseComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
