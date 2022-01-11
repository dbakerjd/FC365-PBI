import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EntityEditFileComponent } from './entity-edit-file.component';

describe('EntityEditFileComponent', () => {
  let component: EntityEditFileComponent;
  let fixture: ComponentFixture<EntityEditFileComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ EntityEditFileComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(EntityEditFileComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
