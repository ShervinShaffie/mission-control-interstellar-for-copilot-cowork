import { useState, type ReactNode } from "react";
import { DndContext, KeyboardSensor, PointerSensor, closestCenter, pointerWithin, useSensor, useSensors, type CollisionDetection, type DragEndEvent, type KeyboardCoordinateGetter, type UniqueIdentifier } from "@dnd-kit/core";
import { SortableContext, useSortable } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { GripVertical, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { componentCatalog, isComponentId, reorderComponent, type ComponentId, type DashboardPreferences } from "@/lib/dashboard-preferences";
import { cn } from "@/lib/utils";

const keepBoxesInPlace = () => null;
const titleOf = (id: UniqueIdentifier) => componentCatalog.find(item => item.id === id)?.title ?? "Component";
const boxCollisions: CollisionDetection = args => args.pointerCoordinates
  ? pointerWithin({ ...args, droppableContainers: args.droppableContainers.filter(item => item.id !== args.active.id) })
  : closestCenter(args);

function LayoutBox({ id, children, hidden, disabled, edge, remove }: {
  id: ComponentId; children: ReactNode; hidden: boolean; disabled: boolean;
  edge?: "before" | "after"; remove: () => void;
}) {
  const { attributes, listeners, setNodeRef, setActivatorNodeRef, transform, isDragging } = useSortable({
    id, disabled, animateLayoutChanges: () => false, transition: null,
  });
  const item = componentCatalog.find(item => item.id === id);
  return <div ref={setNodeRef} id={`component-${id}`} hidden={hidden} tabIndex={-1}
    role="group" aria-label={item?.title} data-layout-box={id} data-drop-edge={edge}
    style={{ transform: isDragging ? CSS.Translate.toString(transform) : undefined }}
    className={cn("layout-box relative flex min-w-0 scroll-mt-5 flex-col rounded-xl focus-visible:outline-2 focus-visible:outline-offset-4 focus-visible:outline-ring",
      item?.wide ? "lg:col-span-2" : "layout-half", isDragging && "z-20 outline-2 outline-offset-4 outline-primary forced-colors:outline-[Highlight]",
      edge && "outline-2 outline-offset-4 outline-primary forced-colors:outline-[Highlight]")}>
    {edge && <span aria-hidden="true" className={cn("pointer-events-none absolute inset-x-0 border-t-4 border-primary forced-colors:border-[Highlight]", edge === "before" ? "-top-2" : "-bottom-2")} />}
    <div role="group" aria-label={`${item?.title} layout controls`} className="mb-2 flex items-start justify-between gap-2" data-pdf-exclude>
      <Button ref={setActivatorNodeRef} id={`move-component-${id}`} variant="outline" size="icon"
        {...attributes} {...listeners} disabled={disabled} aria-label={`Move ${item?.title}`}
        title={`Move ${item?.title}`} className="size-10 shrink-0 touch-none cursor-grab active:cursor-grabbing">
        <GripVertical aria-hidden="true" />
      </Button>
      <Button variant="outline" size="icon" aria-label={`Remove ${item?.title} from dashboard`}
        title={`Remove ${item?.title} from dashboard`} onClick={remove} className="size-10 shrink-0">
        <X aria-hidden="true" />
      </Button>
    </div>
    {children}
  </div>;
}

export function DashboardLayout({ preferences, update, components, focusOnly, emptyState }: {
  preferences: DashboardPreferences; update: (next: DashboardPreferences) => void;
  components: Record<ComponentId, ReactNode>; focusOnly: boolean; emptyState: ReactNode;
}) {
  const [drag, setDrag] = useState<{ active: ComponentId; over?: ComponentId } | null>(null);
  const [announcement, setAnnouncement] = useState("");
  const visible = preferences.visible;
  // Mixed-width, variable-height boxes can make geometric arrow navigation skip
  // neighbors. Move through saved layout order, centering on each target instead.
  const keyboardCoordinates: KeyboardCoordinateGetter = (event, { currentCoordinates, context }) => {
    const direction = event.code === "ArrowDown" || event.code === "ArrowRight" ? 1
      : event.code === "ArrowUp" || event.code === "ArrowLeft" ? -1 : 0;
    if (!direction) return;
    event.preventDefault();
    const current = context.over?.id ?? context.active?.id;
    if (!isComponentId(current) || !context.collisionRect) return;
    const next = visible[visible.indexOf(current) + direction];
    if (!next) return;
    const target = context.droppableRects.get(next);
    if (!target || context.droppableContainers.get(next)?.disabled) return;
    return {
      x: currentCoordinates.x + target.left + target.width / 2 - context.collisionRect.left - context.collisionRect.width / 2,
      y: currentCoordinates.y + target.top + target.height / 2 - context.collisionRect.top - context.collisionRect.height / 2,
    };
  };
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 8 } }),
    useSensor(KeyboardSensor, { coordinateGetter: keyboardCoordinates, scrollBehavior: "auto" }),
  );
  function placement(active: UniqueIdentifier, over: UniqueIdentifier) {
    return isComponentId(active) && isComponentId(over) && visible.indexOf(active) < visible.indexOf(over) ? "after" : "before";
  }
  function preview(active: UniqueIdentifier, over?: UniqueIdentifier) {
    return over && over !== active
      ? `Drop ${titleOf(active)} ${placement(active, over)} ${titleOf(over)}.`
      : `Moving ${titleOf(active)}. Choose another box; Escape cancels.`;
  }
  function finish({ active, over }: DragEndEvent) {
    setDrag(null);
    if (focusOnly || !isComponentId(active.id) || !over || !isComponentId(over.id)) return;
    const next = reorderComponent(visible, active.id, over.id);
    if (next !== visible) update({ ...preferences, visible: next });
  }
  function remove(id: ComponentId) {
    const index = visible.indexOf(id);
    const next = visible.filter(value => value !== id);
    update({ ...preferences, visible: next });
    setAnnouncement(`${titleOf(id)} removed from your dashboard; add it back from Customize, Removed components. Source data was not deleted.`);
    const focusId = next[Math.min(index, next.length - 1)];
    requestAnimationFrame(() => {
      const target = focusId ? document.getElementById(`move-component-${focusId}`) : null;
      // A disabled handle (one box or focus-only) cannot receive focus.
      if (target instanceof HTMLButtonElement && !target.disabled) target.focus();
      else (focusId ? document.getElementById(`component-${focusId}`) : document.getElementById("customize-components"))?.focus();
    });
  }
  return <DndContext sensors={sensors} collisionDetection={boxCollisions}
    accessibility={{
      restoreFocus: true,
      screenReaderInstructions: { draggable: "To move a box, press Space or Enter on its Move handle. Right or Down chooses the next box in layout order; Left or Up chooses the previous box. Press Space or Enter to drop, or Escape to cancel. Move up and Move down are also available in Customize." },
      announcements: {
        onDragStart: ({ active }) => `Moving ${titleOf(active.id)}. Use arrow keys to choose another box, Space to drop, Escape to cancel.`,
        onDragOver: ({ active, over }) => preview(active.id, over?.id),
        onDragEnd: ({ active, over }) => {
          if (focusOnly || !isComponentId(active.id) || !over || !isComponentId(over.id)) return `${titleOf(active.id)} was not moved.`;
          const next = reorderComponent(visible, active.id, over.id);
          return next === visible ? `${titleOf(active.id)} was not moved.` : `${titleOf(active.id)} moved to position ${next.indexOf(active.id) + 1} of ${next.length}.`;
        },
        onDragCancel: ({ active }) => `Moving ${titleOf(active.id)} cancelled. Your layout is unchanged.`,
      },
    }}
    onDragStart={({ active }) => { if (isComponentId(active.id)) setDrag({ active: active.id }); }}
    onDragOver={({ over }) => setDrag(current => current ? { ...current, over: isComponentId(over?.id) ? over.id : undefined } : null)}
    onDragEnd={finish} onDragCancel={() => setDrag(null)}>
    <div className="space-y-3">
      <p role="status" aria-live="polite" aria-atomic="true" className="sr-only">{announcement}</p>
      <SortableContext items={visible} strategy={keepBoxesInPlace}>
        {visible.length ? <div className="grid items-stretch gap-7 lg:grid-cols-2">
          {visible.map(id => <LayoutBox key={id} id={id} hidden={focusOnly && id !== "today"}
            disabled={focusOnly || visible.length < 2} remove={() => remove(id)}
            edge={drag?.over === id && drag.active !== id ? placement(drag.active, id) : undefined}>
            {components[id]}
          </LayoutBox>)}
        </div> : emptyState}
      </SortableContext>
    </div>
  </DndContext>;
}
