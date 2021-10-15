package ice.data

import scala.collection.{Iterator, mutable}
import scala.collection.mutable.ArrayBuffer

trait SyncBuffer[A] extends mutable.Buffer[A] {
  abstract override def length: Int = synchronized {
    super.length
  }

  abstract override def iterator: Iterator[A] = synchronized {
    super.iterator
  }

  abstract override def apply(i: Int): A = synchronized {
    super.apply(i)
  }

  override def applyOrElse[A1 <: Int, B1 >: A](x: A1, default: A1 => B1): B1 = synchronized {
    super.applyOrElse(x, default)
  }

  abstract override def addOne(elem: A): SyncBuffer.this.type = synchronized {
    super.addOne(elem)
  }

  override def addAll(xs: IterableOnce[A]): SyncBuffer.this.type = synchronized {
    super.addAll(xs)
  }

  override def subtractOne(x: A): SyncBuffer.this.type = synchronized {
    super.subtractOne(x)
  }

  override def subtractAll(xs: IterableOnce[A]): SyncBuffer.this.type = synchronized {
    super.subtractAll(xs)
  }

  abstract override def prepend(elem: A): SyncBuffer.this.type = synchronized {
    super.prepend(elem)
  }

  override def prependAll(elems: IterableOnce[A]): SyncBuffer.this.type = synchronized {
    super.prependAll(elems)
  }

  abstract override def insert(idx: Int, elem: A): Unit = synchronized {
    super.insert(idx, elem)
  }

  abstract override def insertAll(idx: Int, elems: IterableOnce[A]): Unit = synchronized {
    super.insertAll(idx, elems)
  }

  abstract override def update(idx: Int, elem: A): Unit = synchronized {
    super.update(idx, elem)
  }

  abstract override def remove(idx: Int): A = synchronized {
    super.remove(idx)
  }

  abstract override def remove(idx: Int, count: Int): Unit = synchronized {
    super.remove(idx, count)
  }

  abstract override def clear(): Unit = synchronized {
    super.clear()
  }

  override def hashCode(): Int = synchronized {
    super.hashCode()
  }

  override def trimStart(n: Int): Unit = synchronized {
    super.trimStart(n)
  }

  override def trimEnd(n: Int): Unit = synchronized {
    super.trimEnd(n)
  }
}

class SyncArrayBuffer[A] extends mutable.ArrayBuffer[A] with SyncBuffer[A] {
  override def clone(): ArrayBuffer[A] = synchronized {
    super.clone()
  }
}
