import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExternalApproveModelComponent } from './external-approve-model.component';

describe('ExternalApproveModelComponent', () => {
  let component: ExternalApproveModelComponent;
  let fixture: ComponentFixture<ExternalApproveModelComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExternalApproveModelComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExternalApproveModelComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
