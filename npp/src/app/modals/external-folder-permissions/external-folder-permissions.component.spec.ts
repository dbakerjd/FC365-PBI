import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExternalFolderPermissionsComponent } from './external-folder-permissions.component';

describe('ExternalFolderPermissionsComponent', () => {
  let component: ExternalFolderPermissionsComponent;
  let fixture: ComponentFixture<ExternalFolderPermissionsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExternalFolderPermissionsComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExternalFolderPermissionsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
